import pdfplumber
import re
from threading import Thread
import datetime
import numpy as np
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import pandas as pd
import os
import unicodedata
from PyPDF2 import PdfReader
import openpyxl

# --- Helper function to get plain text ---
def get_plain_text_pypdf2(pdf_path):
    try:
        reader = PdfReader(pdf_path); text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text: text += page_text + "\n"
        return text
    except Exception as e: print(f"Error extracting text with PyPDF2: {e}"); return ""

# --- Core extraction logic for a SINGLE shipment's text chunks ---
def extract_single_shipment(pdfplumber_chunk, plain_chunk):
    results = {
        'sender_info': {'raw_address': ''}, 'receiver_info': {'raw_address': ''},
        'shipment_details': {}, 'billing_charges': [], 'general_comments': [],
        'weights': {}, 'invoice_history': [], 'billing_ref_codes': []
    }
    if not pdfplumber_chunk and not plain_chunk: return None

    # --- Use pdfplumber_chunk for layout-dependent sections ---
    if pdfplumber_chunk:
        # Shipment Details & Billing Ref Codes
        # Initial extraction using the simpler regex (might include extra text)
        sd = {}
        m_date = re.search(r'Ship Date / Date d’expédition\s*:\s*([^\n]+)', pdfplumber_chunk, re.IGNORECASE)
        if m_date: sd['ship_date'] = m_date.group(1).strip()
        m_service = re.search(r'Service Name / Nom du service\s*:\s*([^\n]+)', pdfplumber_chunk, re.IGNORECASE)
        if m_service: sd['service_name'] = m_service.group(1).strip()
        m_track = re.search(r'Tracking # / # de suivi\s*:\s*([^\n]+)', pdfplumber_chunk, re.IGNORECASE)
        if m_track: sd['tracking_number'] = m_track.group(1).strip()

        # --- START: Post-processing cleanup ---
        if sd:
            pattern_to_remove = r'\s*Billing Ref\. Code.*' # Pattern to remove
            for key in ['ship_date', 'service_name', 'tracking_number']:
                if key in sd and sd[key]: # Check if key exists and value is not None/empty
                     # Use re.sub to remove the unwanted pattern from the string
                     cleaned_value = re.sub(pattern_to_remove, '', sd[key], flags=re.IGNORECASE)
                     # Update the dictionary with the cleaned value
                     sd[key] = cleaned_value.strip() # Remove trailing whitespace
            results['shipment_details'] = sd
        # --- END: Post-processing cleanup ---


        ref_code_matches = re.findall(r'Billing Ref\. Code #\d+\s*/.*?:\s*(\S+)', pdfplumber_chunk, re.IGNORECASE)
        if ref_code_matches: results['billing_ref_codes'] = [code.strip() for code in ref_code_matches]

        # Weight Information (remains the same)
        weight_match = re.search(r'Total # of Packages / # total de colis\s*:\s*(\d+).*?'r'Total Shipment Weight / Poids total de l’expédition\s*:\s*([\d.]+)\s*(lbs|kg).*?'r'Total Billed Weight / Poids total facturé\s*:\s*([\d.]+)\s*(lbs|kg)', pdfplumber_chunk, re.DOTALL | re.IGNORECASE)
        if weight_match: results['weights'] = { 'total_packages': int(weight_match.group(1)), 'shipment_weight': f"{weight_match.group(2).strip()} {weight_match.group(3).strip()}", 'billed_weight': f"{weight_match.group(4).strip()} {weight_match.group(5).strip()}"}

        # Invoice History (remains the same)
        invoice_history_matches = re.findall(r'Invoice/Facture\s+#(\d+)\s+-\s+(FC\d+)\s*:\s+\$?([\d.]+)', pdfplumber_chunk, re.IGNORECASE)
        for inv in invoice_history_matches: results['invoice_history'].append({'invoice_number': inv[0].strip(),'reference': inv[1].strip(),'amount': float(inv[2].strip())})

    # --- Use plain_chunk for linearized sections ---
    # (Address and Billing logic remains the same as the previous corrected version)
    if plain_chunk:
        # Addresses
        sender_lines = []; receiver_lines = []
        address_match = re.search(r'(?:Sender / Expéditeur :.*?Receiver / Récepteur :)\n(.*?)\nShipment Information', plain_chunk, re.DOTALL | re.IGNORECASE)
        if address_match:
            address_block = address_match.group(1).strip(); block_lines = address_block.split('\n'); split_point = -1
            for data_line in block_lines[:5]:
                line_content = data_line.strip();
                if len(line_content) > 15: # Check reasonably long lines
                    # Look for 2+ spaces followed by an uppercase letter or digit, not too close to start
                    match = re.search(r'\s{2,}(?=[A-Z0-9])', line_content)
                    if match and match.start() > 5: # Ensure gap isn't right at the beginning
                        potential_split = match.start() + 1 # Assign ONLY if match is valid

                        # --- Moved Check Inside ---
                        # Check if potential split creates two reasonable parts AFTER assigning it
                        if potential_split > 3 and (len(line_content) - potential_split) > 3:
                             split_point = potential_split # Assign the final split_point
                             # print(f"DEBUG: Found split point {split_point} on line: {line_content}") # Optional Debug
                             break # Exit the loop once a suitable split point is found
           
            for line in block_lines:
                line_stripped = line.strip();
                if not line_stripped: continue
                if split_point != -1 and len(line) >= split_point: 
                    sender_part = line[:split_point].strip(); receiver_part = line[split_point:].strip();
                    if sender_part:
                        sender_lines.append(sender_part);
                    if receiver_part: 
                        receiver_lines.append(receiver_part)
                elif split_point != -1 and line_stripped: sender_lines.append(line_stripped)
                else: sender_lines.append(line_stripped)
            results['sender_info']['raw_address'] = "\n".join(sender_lines); results['receiver_info']['raw_address'] = "\n".join(receiver_lines)
            if split_point == -1 and sender_lines: print("Warning: Could not determine address split point...") # Simplified warning
        else: results['sender_info']['raw_address'] = "Address block not found"; results['receiver_info']['raw_address'] = "Address block not found"

        # Billing Charges
        billing_charges = []
        billing_section_match = re.search(r'Billing Details.*?facturation\n(.*?)(?=\n(?:Original Dimensions|Shipment Invoice History|Sender / Expéditeur :))', plain_chunk, re.DOTALL | re.IGNORECASE)
        if billing_section_match:
            billing_lines = billing_section_match.group(1).strip().split('\n')
            current_charge_type_lines = []; current_comments = []; current_billed_amount = None; current_quoted_amount = None; last_charge_added = False
            def finalize_charge(type_lines, quoted, billed, comments_list, local_billing_charges):
                if not type_lines or billed is None: return None
                charge_type_raw = ' '.join(type_lines).replace('supplémentaires', '').strip(); charge_type_key = re.split(r'\s*/', charge_type_raw, 1)[0].strip()
                quoted_val = quoted if quoted is not None else None; charge = {'charge_type': charge_type_key, 'quoted': quoted_val, 'billed': billed, 'comments': ' '.join(comments_list).strip()}; local_billing_charges.append(charge); return charge
            for i, line in enumerate(billing_lines):
                line = line.strip(); last_charge_added = False
                if not line or any(hdr in line for hdr in ['Quoted Charges', 'Billed Charges', 'Comments / Commentaires']) or 'Total:' in line:
                    if current_charge_type_lines and current_billed_amount is not None: finalize_charge(current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments, billing_charges); last_charge_added = True
                    current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments = [], None, None, []; continue
                is_type_line = '/' in line; is_potential_amount_line = bool(re.search(r'\$?[\d.]+', line)); is_amount_line = is_potential_amount_line and not is_type_line; is_comment_end = line.endswith('.')
                if is_type_line:
                    if current_charge_type_lines and (current_billed_amount is not None or current_quoted_amount is not None): finalize_charge(current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments, billing_charges); last_charge_added = True; current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments = [line], None, None, []
                    elif current_charge_type_lines: current_charge_type_lines.append(line)
                    else: current_charge_type_lines = [line]; current_comments = []; current_quoted_amount = None; current_billed_amount = None
                elif is_amount_line:
                    amounts_found = re.findall(r'\$?([\d.]+)', line); comment_part = re.sub(r'\$?[\d.]+\s*', '', line).strip(); temp_quoted = None; temp_billed = None; valid_amounts_parsed = False
                    try:
                        if len(amounts_found) == 2: temp_quoted = float(amounts_found[0]); temp_billed = float(amounts_found[1]); valid_amounts_parsed = True
                        elif len(amounts_found) == 1: temp_quoted = None; temp_billed = float(amounts_found[0]); valid_amounts_parsed = True
                    except ValueError: print(f"Warning: Could not convert amount to float on line: {line}")
                    if valid_amounts_parsed:
                        if current_billed_amount is not None or current_quoted_amount is not None: print(f"Warning: Overwriting existing amount data for {current_charge_type_lines} with line '{line}'.");
                        current_quoted_amount = temp_quoted; current_billed_amount = temp_billed
                        if comment_part: current_comments.append(comment_part)
                        if comment_part and comment_part.endswith('.'):
                            if current_charge_type_lines and current_billed_amount is not None: finalize_charge(current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments, billing_charges); last_charge_added = True; current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments = [], None, None, []
                    else: current_comments.append(line) # Treat as comment if amount parsing failed
                    if line.endswith('.'): # Still check if it ends comments even if amounts failed
                        if current_charge_type_lines and current_billed_amount is not None: finalize_charge(current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments, billing_charges); last_charge_added = True; current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments = [], None, None, []
                else:
                    current_comments.append(line)
                    if is_comment_end:
                        if current_charge_type_lines and current_billed_amount is not None: finalize_charge(current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments, billing_charges); last_charge_added = True; current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments = [], None, None, []
            if not last_charge_added and current_charge_type_lines and current_billed_amount is not None: finalize_charge(current_charge_type_lines, current_quoted_amount, current_billed_amount, current_comments, billing_charges)
            results['billing_charges'] = billing_charges
        else: results['billing_charges'] = []


    if not results.get('shipment_details') and not results.get('billing_charges') and not results.get('weights'): print("Warning: Minimal data extracted...")
    return results

# --- Main Orchestrator Function ---
# (process_multipage_freightcom_pdf function remains the same)
def process_multipage_freightcom_pdf(pdf_path):
    print(f"Processing PDF: {pdf_path}")
    pdfplumber_text_full = ""
    try:
        with pdfplumber.open(pdf_path) as pdf: pdfplumber_text_full = "\n".join([page.extract_text(x_tolerance=1, y_tolerance=1) or "" for page in pdf.pages])
    except Exception as e: print(f"Error extracting text with pdfplumber: {e}")
    plain_text_full = get_plain_text_pypdf2(pdf_path)
    if not plain_text_full: print("Critical Error: Plain text extraction failed."); return []
    if not pdfplumber_text_full: print("Warning: pdfplumber text extraction failed.")

    delimiter = "Sender / Expéditeur :"; plain_text_markers = [m.start() for m in re.finditer(re.escape(delimiter), plain_text_full, re.IGNORECASE)]
    if not plain_text_markers: print(f"Warning: Shipment start marker '{delimiter}' not found."); return []
    pdfplumber_markers = []
    if pdfplumber_text_full: pdfplumber_markers = [m.start() for m in re.finditer(re.escape(delimiter), pdfplumber_text_full, re.IGNORECASE)]
    if len(plain_text_markers) != len(pdfplumber_markers) and pdfplumber_text_full: print(f"Warning: Mismatch in marker counts plain({len(plain_text_markers)}) vs pdfplumber({len(pdfplumber_markers)}).")

    all_shipment_data = []; print(f"Found {len(plain_text_markers)} potential shipment start markers.")
    for i in range(len(plain_text_markers)):
        print(f"\n--- Processing Shipment Record {i+1} ---")
        chunk_start_plain = plain_text_markers[i]; chunk_end_plain = plain_text_markers[i+1] if (i + 1) < len(plain_text_markers) else len(plain_text_full)
        current_plain_chunk = plain_text_full[chunk_start_plain:chunk_end_plain]
        current_pdfplumber_chunk = ""
        if pdfplumber_text_full and i < len(pdfplumber_markers):
             chunk_start_pdfplumber = pdfplumber_markers[i]; chunk_end_pdfplumber = pdfplumber_markers[i+1] if (i + 1) < len(pdfplumber_markers) else len(pdfplumber_text_full)
             current_pdfplumber_chunk = pdfplumber_text_full[chunk_start_pdfplumber:chunk_end_pdfplumber]
        elif pdfplumber_text_full: print(f"Warning: No corresponding pdfplumber marker found for shipment {i+1}.")
        shipment_data = extract_single_shipment(current_pdfplumber_chunk, current_plain_chunk)
        if shipment_data: all_shipment_data.append(shipment_data)
        else: print(f"Notice: No significant data extracted for shipment record {i+1}, skipping.")

    print(f"\n--- Finished processing. Extracted data for {len(all_shipment_data)} shipment records. ---")
    return all_shipment_data

# --- Function to Sanitize Column Names ---
# (sanitize_column_name function remains the same)
def sanitize_column_name(name):
    nfkd_form = unicodedata.normalize('NFKD', name); only_ascii = nfkd_form.encode('ASCII', 'ignore').decode('ASCII')
    safe_name = re.sub(r'[^\w]+', '_', only_ascii); safe_name = safe_name.lower().strip('_')
    if not safe_name: return "charge_unknown"
    return f"charge_{safe_name}"

# --- Updated Function to Flatten Data (Pivoting Charges with Quoted/Billed) and Export ---
# (flatten_and_export_data_pivoted function remains the same - logic for quoted columns was already present)
def flatten_and_export_data_pivoted(data_list):
    if not data_list: print("No data extracted to export."); return

    all_flat_data = []
    for shipment_data in data_list:
        flat_data = {}

        # Shipment Details
        details = shipment_data.get('shipment_details', {})
        flat_data['shipment_ship_date'] = details.get('ship_date')
        flat_data['shipment_waybill_id'] = details.get('tracking_number') # Rename
        flat_data['shipment_service_name'] = details.get('service_name')
        flat_data['shipment_billing_ref_codes'] = '|'.join(shipment_data.get('billing_ref_codes', []))

        # Weights
        weights = shipment_data.get('weights', {}); flat_data['weights_total_packages'] = weights.get('total_packages'); flat_data['weights_shipment_weight'] = weights.get('shipment_weight'); flat_data['weights_billed_weight'] = weights.get('billed_weight')

        # Addresses
        sender = shipment_data.get('sender_info', {}); flat_data['address'] = sender.get('raw_address');
        # Invoice History
        history = shipment_data.get('invoice_history', []); flat_data['history_invoice_numbers'] = '|'.join([h.get('invoice_number', '') for h in history]); flat_data['history_references'] = '|'.join([h.get('reference', '') for h in history]); flat_data['history_amounts'] = '|'.join([str(h.get('amount', '')) for h in history])

        # --- Process Billing Charges - PIVOT with Quoted/Billed ---
        charges = shipment_data.get('billing_charges', [])
        all_charge_comments_list = []

        for charge in charges:
            charge_type = charge.get('charge_type'); billed_amount = charge.get('billed'); quoted_amount = charge.get('quoted'); comment = charge.get('comments', '').strip()
            if charge_type:
                base_col_name = sanitize_column_name(charge_type).replace('charge_', '', 1)
                col_name_billed = f"charge_{base_col_name}_billed"; col_name_quoted = f"charge_{base_col_name}_quoted"
                # Assign Billed
                if col_name_billed in flat_data:
                    try: current_val = flat_data[col_name_billed] if pd.notna(flat_data[col_name_billed]) else 0; new_val = billed_amount if pd.notna(billed_amount) else 0; flat_data[col_name_billed] = current_val + new_val
                    except TypeError: flat_data[col_name_billed] = billed_amount if pd.notna(billed_amount) else None
                    print(f"Warning: Duplicate charge type '{charge_type}' -> column '{col_name_billed}'. Amounts summed.")
                else: flat_data[col_name_billed] = billed_amount if pd.notna(billed_amount) else None
                # Assign Quoted (only if not None)
                if quoted_amount is not None:
                    if col_name_quoted in flat_data:
                        try: current_val = flat_data[col_name_quoted] if pd.notna(flat_data[col_name_quoted]) else 0; new_val = quoted_amount if pd.notna(quoted_amount) else 0; flat_data[col_name_quoted] = current_val + new_val
                        except TypeError: flat_data[col_name_quoted] = quoted_amount if pd.notna(quoted_amount) else None
                        print(f"Warning: Duplicate charge type '{charge_type}' -> column '{col_name_quoted}'. Amounts summed.")
                    else: flat_data[col_name_quoted] = quoted_amount if pd.notna(quoted_amount) else None
            if comment: all_charge_comments_list.append(comment)

        flat_data['all_charge_comments'] = '|'.join(all_charge_comments_list)
        all_flat_data.append(flat_data)

    df = pd.DataFrame(all_flat_data)

    # --- Define Column Order ---
    core_columns_start = ['waybill', 'shipment_ship_date', 'shipment_service_name', 'shipment_billing_ref_codes', 'weights_total_packages', 'weights_shipment_weight', 'weights_billed_weight', 'history_invoice_numbers', 'history_references', 'history_amounts']
    charge_columns = sorted([col for col in df.columns if col.startswith('charge_') and ('_billed' in col or '_quoted' in col)])
    core_columns_end = ['address', 'all_charge_comments', ]
    final_column_order = core_columns_start + charge_columns + core_columns_end
    existing_columns_in_order = [col for col in final_column_order if col in df.columns]
    final_columns_to_use = existing_columns_in_order + [col for col in df.columns if col not in existing_columns_in_order]
    df = df[final_columns_to_use]
    return df


# --- Test Execution ---
if __name__ == "__main__":
    # pdf_file = "d3.pdf" # Test file 1
    pdf_file = "detail.pdf" # Test file 2 (assuming you saved the snippet)
    # pdf_file = "path/to/your/multipage_freightcom.pdf" # Use actual multi-page file

    try:
        extracted_data_list = process_multipage_freightcom_pdf(pdf_file)
        flatten_and_export_data_pivoted(extracted_data_list)
    except FileNotFoundError: print(f"Error: Input PDF file not found at {pdf_file}")
    except Exception as e: print(f"An unexpected error occurred during the process: {e}"); import traceback; traceback.print_exc()





# Define the BRAND logic function (similar to JS)
def determine_brand(row):
    client_name = row.get('ClientName', '') # Use .get for safety
    customer_number = row.get('CustomerNumber', '')

    # Ensure customer_number is treated as string for comparison if needed
    customer_number_str = str(customer_number) if pd.notna(customer_number) else ""

    if client_name == "Maddle Boards":
        return "Maddle"
    elif client_name == "Montreal Weights / Ascend":
        if customer_number_str == "SHOPIFYAI1":
            return "Ascend"
        elif customer_number_str == "SHOPIFYAI2":
            return "Nordik"
        elif customer_number_str == "SHOPIFYAI":
            return "Montreal Weights"
    # Original JS had a check for !customerNumber - translate to check if it's None/NaN/empty
    elif not customer_number_str:
         return "To Review"
    return "" # Default empty string


# Define the comparison logic function
def perform_comparison(extracted_df, compare_csv_path, output_folder, status_callback):
    """
    Performs the comparison logic based on the JS example.

    Args:
        extracted_df (pd.DataFrame): DataFrame from PDF extraction.
        compare_csv_path (str): Path to the CSV file to compare against.
        output_folder (str): Folder to save the merged output file.
        status_callback (function): Function to update the GUI status label.
    """
    try:
        status_callback("Loading comparison CSV...")
        compare_df = pd.read_csv(compare_csv_path)

        if extracted_df is None or extracted_df.empty:
             raise ValueError("Extracted data is empty. Cannot perform comparison.")
        if compare_df.empty:
            raise ValueError("Comparison CSV is empty. Cannot perform comparison.")

        # --- 1. Add BRAND column to compare_df ---
        status_callback("Adding BRAND column...")
        if 'ClientName' not in compare_df.columns or 'CustomerNumber' not in compare_df.columns:
             # Warn but continue, brand column will be mostly empty or default
             messagebox.showwarning("Missing Columns",
                                   "Comparison CSV missing 'ClientName' or 'CustomerNumber'. 'BRAND' column may be incomplete.")
             compare_df['BRAND'] = "" # Add empty column if source columns are missing
        else:
            # Ensure CustomerNumber is treated as string for comparison consistency
             compare_df['CustomerNumber'] = compare_df['CustomerNumber'].astype(str)
             compare_df['BRAND'] = compare_df.apply(determine_brand, axis=1)


        # --- 2. Calculate charges_total in extracted_df ---
        status_callback("Calculating charges total...")
        # Make column names lowercase for case-insensitive matching
        extracted_df.columns = map(str.lower, extracted_df.columns)
        if 'total' in extracted_df.columns:
            try:
                total_index = extracted_df.columns.get_loc('total')
                # Select columns after 'total', exclude those containing 'st'
                relevant_columns = [
                    col for col in extracted_df.columns[total_index + 1:]
                    if 'st' not in col.lower()
                ]
                if relevant_columns:
                     # Ensure relevant columns are numeric, coercing errors to NaN
                    for col in relevant_columns:
                        extracted_df[col] = pd.to_numeric(extracted_df[col], errors='coerce')
                    # Sum, treating NaNs as 0 for the sum
                    extracted_df['charges_total'] = extracted_df[relevant_columns].sum(axis=1, skipna=True)
                else:
                    extracted_df['charges_total'] = 0 # Add column even if no relevant charges found
                    print("No non-'st' columns found after 'total' to sum for 'charges_total'.")
            except Exception as e:
                 print(f"Error calculating charges_total: {e}. Skipping calculation.")
                 extracted_df['charges_total'] = np.nan # Indicate calculation failed


        # --- 3. Reorder 'st' columns in extracted_df ---
        status_callback("Reordering columns...")
        if 'total' in extracted_df.columns:
            st_columns = [col for col in extracted_df.columns if 'st' in col.lower()]
            other_columns = [col for col in extracted_df.columns if col not in st_columns]

            try:
                total_index_new = other_columns.index('total') # Find index in the *filtered* list
                reordered_cols = (
                    other_columns[:total_index_new] +
                    st_columns +
                    other_columns[total_index_new:]
                )
                extracted_df = extracted_df[reordered_cols]
            except ValueError:
                 print("'total' column not found after filtering 'st' columns. Skipping reorder.")
            except Exception as e:
                print(f"Error reordering columns: {e}. Skipping reorder.")
        else:
            print("'total' column not found in extracted data. Skipping column reordering.")


        # --- 4. Add waybill to compare_df for merging ---
        status_callback("Matching waybills...")
        if 'waybill' not in extracted_df.columns:
            raise ValueError("Critical Error: 'waybill' column missing in extracted data.")
        if 'TrackingNumber' not in compare_df.columns:
            raise ValueError("Critical Error: 'TrackingNumber' column missing in comparison CSV.")

        # Ensure waybill in extracted_df is usable (e.g., string, not NaN)
        extracted_waybills = set(extracted_df['waybill'].dropna().astype(str))
        if not extracted_waybills:
             messagebox.showwarning("Waybill Matching", "No valid 'waybill's found in extracted data.")
             compare_df['waybill_match'] = pd.NA # Use pandas NA for clarity
        else:
            # Define the matching function (handles potential errors)
            def find_match(tracking_num):
                if pd.isna(tracking_num):
                    return pd.NA # Explicitly handle NaN tracking numbers
                try:
                    # Ensure tracking_num is string for 'in' check
                    tracking_num_str = str(tracking_num)
                    # Find the *first* matching extracted waybill contained within the tracking number
                    for wb in extracted_waybills:
                        if wb in tracking_num_str:
                            return wb # Return the matching waybill from extracted_df
                    return pd.NA # No match found
                except Exception as e:
                    # print(f"Error matching tracking number {tracking_num}: {e}") # Optional: verbose logging
                    return pd.NA # Error during matching

            # Apply the matching function
            compare_df['waybill_match'] = compare_df['TrackingNumber'].apply(find_match)

        # --- 5. Filter compare_df for rows with matches ---
        status_callback("Filtering matches...")
        # Use the *newly created* waybill_match column for filtering and merging
        filtered_compare_df = compare_df.dropna(subset=['waybill_match'])

        if filtered_compare_df.empty:
            messagebox.showwarning("No Matches", "No matching waybills found between the two files.")
            status_callback("Comparison failed: No matches found.")
            return None # Indicate no merge happened

        # --- 6. Merge the DataFrames ---
        status_callback("Merging data...")
        # Merge using the matched waybill ID from compare_df and the original waybill from extracted_df
        # We rename the match column in filtered_compare_df temporarily for the merge key
        merged_df = pd.merge(
            extracted_df,
            filtered_compare_df,
            left_on='waybill',          # key from extracted_df
            right_on='waybill_match', # key from compare_df (that we created)
            how='inner'                    # Keep only matching rows
        )

        # Optional: Drop the redundant match column if desired
        if 'waybill_match' in merged_df.columns:
             merged_df = merged_df.drop(columns=['waybill_match'])


        # --- 7. Drop empty columns ---
        status_callback("Cleaning up merged data...")
        # Drop columns where ALL values are NaN or None
        merged_df = merged_df.dropna(axis=1, how='all')


        # --- 8. Save the result ---
        status_callback("Saving merged file...")
        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        # Try saving as Excel first
        excel_filename = os.path.join(output_folder, f"merged-{timestamp}.xlsx")
        csv_filename = os.path.join(output_folder, f"merged-{timestamp}.csv") # Fallback path

        try:
            merged_df.to_excel(excel_filename, index=False, sheet_name="Merged Data")
            final_path = excel_filename
            status_callback(f"Success! Merged file saved to {final_path}")
            messagebox.showinfo("Comparison Complete", f"Merged file saved as:\n{final_path}")
        except ImportError:
            print("`openpyxl` not installed. Falling back to CSV.")
            print("Install with: pip install openpyxl")
            merged_df.to_csv(csv_filename, index=False)
            final_path = csv_filename
            status_callback(f"Success! Merged file saved to {final_path} (CSV format)")
            messagebox.showinfo("Comparison Complete", f"Merged file saved as (CSV fallback):\n{final_path}")
        except Exception as excel_error:
            print(f"Error writing Excel file: {excel_error}")
            print("Falling back to saving as CSV.")
            try:
                merged_df.to_csv(csv_filename, index=False)
                final_path = csv_filename
                status_callback(f"Success! Merged file saved to {final_path} (CSV format)")
                messagebox.showinfo("Comparison Complete", f"Merged file saved as (CSV fallback):\n{final_path}")
            except Exception as csv_error:
                print(f"FATAL: Error writing fallback CSV file: {csv_error}")
                status_callback("Error: Could not save merged file.")
                messagebox.showerror("Save Error", f"Could not save the merged file as Excel or CSV.\n\nExcel Error: {excel_error}\n\nCSV Error: {csv_error}")
                final_path = None

        return final_path # Return the path where it was saved (or None if failed)

    except FileNotFoundError as e:
        status_callback("Error: Comparison file not found.")
        messagebox.showerror("File Not Found", f"Could not find the comparison file:\n{e.filename}")
        print(traceback.format_exc())
        return None
    except ValueError as e: # Catch specific errors raised in the logic
        status_callback(f"Error: {e}")
        messagebox.showerror("Comparison Error", str(e))
        print(traceback.format_exc())
        return None
    except KeyError as e: # Catch missing columns
         col_name = e.args[0]
         status_callback(f"Error: Required column '{col_name}' not found.")
         messagebox.showerror("Missing Column", f"A required column ('{col_name}') was not found in one of the files.")
         print(traceback.format_exc())
         return None
    except Exception as e:
        status_callback("Error during comparison.")
        messagebox.showerror("Error", f"An unexpected error occurred during comparison:\n{e}")
        print(traceback.format_exc())
        return None


# Main PDF processing function (runs in thread)
def process_pdfs_thread(input_folder, status_callback, completion_callback):
    """
    Processes PDFs in a folder, aggregates results, and calls back on completion.

    Args:
        input_folder (str): Path to the folder containing PDFs.
        status_callback (function): Function to update the GUI status label.
        completion_callback (function): Function to call when done, passing the DataFrame or None.
    """
    all_dfs = []
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]

    if not pdf_files:
        status_callback("No PDF files found in selected folder")
        completion_callback(None) # Signal completion with no data
        return

    total_files = len(pdf_files)
    status_callback(f"Starting processing for {total_files} PDF(s)...")

    for i, pdf_file in enumerate(pdf_files):
        pdf_path = os.path.join(input_folder, pdf_file)
        try:
            status_callback(f"Processing file {i+1}/{total_files}: {pdf_file}")
            df  = flatten_and_export_data_pivoted(process_multipage_freightcom_pdf(pdf_path)) # Call your extraction function
            if df is not None and not df.empty:
                all_dfs.append(df)
        except Exception as e:
            print(f"Error processing {pdf_file}: {str(e)}")
            # Optionally, show a warning for individual file errors
            # messagebox.showwarning("File Error", f"Could not process {pdf_file}:\n{e}")

    if all_dfs:
        try:
            status_callback("Combining extracted data...")
            final_df = pd.concat(all_dfs, ignore_index=True)
            # Basic validation: Ensure 'waybill' exists post-concat
            if 'waybill' not in final_df.columns:
                 print("Warning: 'waybill' column not found after combining PDFs. Comparison might fail.")
                 # Decide if this is a critical error or just a warning
                 # completion_callback(None) # Or pass the df anyway? Depends on requirements.
            status_callback(f"Processing complete. Found data from {len(all_dfs)} PDF(s).")
            completion_callback(final_df) # Pass the combined DataFrame back
        except Exception as e:
            status_callback("Error combining data.")
            messagebox.showerror("Error", f"Failed to combine data: {e}")
            print(traceback.format_exc())
            completion_callback(None) # Signal completion with error
    else:
        status_callback("No data extracted from any PDF files.")
        messagebox.showwarning("No Data", "No valid shipment data was extracted from the PDF files.")
        completion_callback(None) # Signal completion with no data


class PDFApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Processor and Comparer")
        self.geometry("900x600") # Increased size slightly

        self.extracted_df = None # To store the DataFrame after PDF processing

        # --- Input Folder Selection ---
        input_frame = ttk.LabelFrame(self, text="1. Select PDF Input Folder")
        input_frame.pack(padx=10, pady=5, fill="x")

        self.input_folder = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.input_folder, width=60).pack(side=tk.LEFT, padx=5, expand=True, fill="x")
        ttk.Button(input_frame, text="Browse...", command=self.select_input_folder).pack(side=tk.LEFT, padx=5)

        # --- Processing Control ---
        process_frame = ttk.Frame(self)
        process_frame.pack(pady=10)
        self.process_button = ttk.Button(process_frame, text="Process PDFs", command=self.start_pdf_processing, state=tk.DISABLED)
        self.process_button.pack()

        # --- Status Label ---
        self.status_label_var = tk.StringVar(value="Select PDF input folder to begin.")
        status_bar = ttk.Label(self, textvariable=self.status_label_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill="x", pady=(5,0), padx=0)

        # --- Action Buttons (Initially hidden/disabled) ---
        action_frame = ttk.LabelFrame(self, text="2. Choose Action for Extracted Data")
        action_frame.pack(padx=10, pady=10, fill="x")

        self.save_button = ttk.Button(action_frame, text="Save Extracted Data as CSV", command=self.save_extracted_csv, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=15, pady=5)

        self.compare_button = ttk.Button(action_frame, text="Compare with Another CSV", command=self.start_comparison, state=tk.DISABLED)
        self.compare_button.pack(side=tk.LEFT, padx=15, pady=5)

        # Set initial default folder for convenience if possible
        try:
             default_input = os.path.expanduser("~/Downloads") # Common starting point
             if os.path.isdir(default_input):
                 self.input_folder.set(default_input)
                 self.process_button.config(state=tk.NORMAL)
                 self.update_status("Ready to process PDFs.")
        except Exception:
            pass # Ignore if default cannot be set

    def update_status(self, message):
        """Safely updates the status bar from any thread."""
        self.status_label_var.set(message)
        # print(message) # Also print to console for debugging

    def select_input_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing PDFs")
        if folder:
            self.input_folder.set(folder)
            self.process_button.config(state=tk.NORMAL)
            self.update_status("Input folder selected. Click 'Process PDFs'.")
            # Reset state if a new folder is selected
            self.extracted_df = None
            self.save_button.config(state=tk.DISABLED)
            self.compare_button.config(state=tk.DISABLED)

    def start_pdf_processing(self):
        input_folder = self.input_folder.get()
        if not input_folder or not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Please select a valid input folder first.")
            return

        # Disable buttons during processing
        self.process_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        self.compare_button.config(state=tk.DISABLED)
        self.extracted_df = None # Clear previous results

        self.update_status("Starting PDF processing...")

        # Run processing in a separate thread
        process_thread = Thread(
            target=process_pdfs_thread,
            args=(input_folder, self.update_status, self.on_pdf_processing_complete),
            daemon=True
        )
        process_thread.start()

    def on_pdf_processing_complete(self, result_df):
        """Callback function executed when PDF processing thread finishes."""
        # Ensure GUI updates happen in the main thread
        self.process_button.config(state=tk.NORMAL) # Re-enable processing button

        if result_df is not None and not result_df.empty:
            self.extracted_df = result_df
            self.update_status("PDF processing complete. Choose an action.")
            # Enable action buttons
            self.save_button.config(state=tk.NORMAL)
            self.compare_button.config(state=tk.NORMAL)
        else:
            # Keep action buttons disabled if no data was extracted or an error occurred
            self.save_button.config(state=tk.DISABLED)
            self.compare_button.config(state=tk.DISABLED)
            # Status is already set by process_pdfs_thread on error/no data
            if self.status_label_var.get().startswith("Processing") or self.status_label_var.get().startswith("Starting"):
                 self.update_status("PDF processing finished, but no data was extracted.")


    def save_extracted_csv(self):
        if self.extracted_df is None or self.extracted_df.empty:
            messagebox.showerror("Error", "No extracted data available to save.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Extracted Data",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            initialfile="extracted_shipments.csv"
        )

        if file_path:
            try:
                self.extracted_df.to_csv(file_path, index=False)
                self.update_status(f"Extracted data saved to {os.path.basename(file_path)}")
                messagebox.showinfo("Save Successful", f"Data saved to:\n{file_path}")
            except Exception as e:
                self.update_status("Error saving extracted data.")
                messagebox.showerror("Save Error", f"Could not save the file:\n{e}")
                print(traceback.format_exc())

    def start_comparison(self):
        if self.extracted_df is None or self.extracted_df.empty:
            messagebox.showerror("Error", "No extracted data available for comparison.")
            return

        # 1. Ask for the comparison CSV
        compare_file = filedialog.askopenfilename(
            title="Select CSV File to Compare Against",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if not compare_file:
            self.update_status("Comparison cancelled (no comparison file selected).")
            return # User cancelled

        # 2. Ask for the output folder (where merged file will be saved)
        # Suggest saving in the same directory as the comparison file initially
        initial_dir = os.path.dirname(compare_file)
        output_folder = filedialog.askdirectory(
             title="Select Folder to Save Merged Output File",
             initialdir=initial_dir
        )

        if not output_folder:
            self.update_status("Comparison cancelled (no output folder selected).")
            return # User cancelled


        # Disable buttons during comparison
        self.process_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        self.compare_button.config(state=tk.DISABLED)
        self.update_status("Starting comparison...")

        # Run comparison in a thread to keep GUI responsive
        comparison_thread = Thread(
            target=self.run_comparison_thread,
             args=(self.extracted_df.copy(), compare_file, output_folder), # Pass a copy of df
            daemon=True
        )
        comparison_thread.start()

    def run_comparison_thread(self, extracted_df_copy, compare_path, output_dir):
         """Wrapper to run perform_comparison and handle GUI updates."""
         try:
             perform_comparison(extracted_df_copy, compare_path, output_dir, self.update_status)
         finally:
             # Re-enable buttons regardless of success or failure
             # Make sure this runs in the main thread if necessary, but should be ok for state changes
             self.process_button.config(state=tk.NORMAL if self.input_folder.get() else tk.DISABLED) # Only enable if input is set
             self.save_button.config(state=tk.NORMAL if self.extracted_df is not None else tk.DISABLED) # Only enable if data exists
             self.compare_button.config(state=tk.NORMAL if self.extracted_df is not None else tk.DISABLED) # Only enable if data exists
             # Final status is set within perform_comparison



if __name__ == "__main__":
    # Ensure necessary libraries are available
    try:
        import pandas
        import numpy
    except ImportError as e:
        messagebox.showerror("Missing Library", f"Required library not found: {e.name}\nPlease install it (e.g., pip install {e.name})")
        exit()

    app = PDFApp()
    app.mainloop()