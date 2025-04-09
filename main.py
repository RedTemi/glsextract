import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np # For vectorized operations if needed
import os
import re
import pdfplumber
from threading import Thread
import datetime
import traceback # For detailed error logging
import openpyxl

# *** Include or import your actual extract_shipments function here ***
# Example (using the placeholder defined above):
# from your_extraction_module import extract_shipments
# Or paste the function definition directly.


def extract_shipments(pdf_path):
    shipments = []
    current_shipment = None
    TERMINATION_PHRASES = {
        "important notice", 
        "invoice summary",
        "for customers paying by direct deposit",
        "kindly send your remittance advice"
    }

    # Regex pattern updates
    waybill_regex = re.compile(
        r'^([A-Z]\d{8})\s+'          # Waybill ID (group 1)
        r'(\d{2}\s\d{2}\s\d{4})\s+'  # Date (group 2)
        r'.*?'                       # Ignore intermediate fields
        r'(\d+\.\d+)\s+'             # Weight (group 3)
        r'(lb|kg)'                   # Unit (group 4)
        r'(?:\s*(.*?)\s+)?base\s+'   # Optional description (group 5)
        r'(\d+\.\d+)',               # Base amount (group 6)
        re.IGNORECASE
    )


    charge_patterns = [
        (r'\bweight\b', 'weight_charge'), (r'carbon\s*surch?r[gq]?\.?', 'carbon surchrg.'),
        (r'\bfuel\b', 'fuel'), (r'2nd\s*delivery', '2nd delivery'),
        (r'adrs\s*correction', 'adrs correction'), (r'ps:\s*max\s*limits', 'ps: max limits'),
        (r'non.?conveyable', 'non-conveyable'), (r'over\s*36\s*inches', 'over 36 inches'),
        (r'over\s*44\s*inches', 'over 44 inches'), (r'over\s*max\s*limits', 'over max limits'),
        (r'overweight\s*\(pc\)', 'overweight (pc)'), (r'overweight\s*\(sh\)', 'overweight (sh)'),
        (r'HST\s*(NB|NFL|NS|ON|PE)?', lambda m: f'HST {m.group(1)}' if m.group(1) else 'HST'),
        (r'\bcredit\b', 'credit'), (r'\bzone\b', 'zone'), (r'GST', 'GST'), (r'QST', 'QST')
    ]

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            for line in lines:
                if any(phrase in line.lower() for phrase in TERMINATION_PHRASES):
                    if current_shipment:
                        shipments.append(current_shipment)
                        current_shipment = None
                    break  # Exit line processing for this page
                               
                waybill_match = waybill_regex.match(line)
                if waybill_match:
                    if current_shipment and 'credit' not in current_shipment['charges']:
                        shipments.append(current_shipment)
                    
                    # Handle description (group 5 might be None)
                    description = waybill_match.group(5).strip().lower() if waybill_match.group(5) else ''
                    
                    if description == "adjustment":
                        current_shipment = None
                        continue
                        
                    current_shipment = {
                        'waybill': waybill_match.group(1),
                        'date': '-'.join(waybill_match.group(2).split()),
                        'weight': float(waybill_match.group(3)),
                        'unit': waybill_match.group(4),
                        'description': description,
                        'total': None,
                        'base': float(waybill_match.group(6)),
                        'charges': {}
                    }
                    continue
                
                if current_shipment:
                   # Charge processing logic remains the same...
                    charge_matches = re.findall(r'([A-Za-z][\w\s\-\(\):\.%]+?)\s+(\d+\.\d+)', line)
                    for charge_raw, amount_str in charge_matches:
                        charge_name = None
                        charge_raw_cleaned = charge_raw.strip()
                        amount = float(amount_str)
                        for pattern, handler in charge_patterns:
                            match = re.search(pattern, charge_raw_cleaned, re.IGNORECASE)
                            if match:
                                charge_name = handler(match) if callable(handler) else handler
                                break
                        if charge_name:
                           current_shipment['charges'][charge_name] = amount

                    total_match = re.search(r'(\d+\.\d+)\s*$', line)
                    if total_match:
                         potential_total = float(total_match.group(1))
                         if potential_total >= current_shipment.get('base', 0):
                              if current_shipment.get('total') is None:
                                  current_shipment['total'] = potential_total


    # Append the last shipment...
    if current_shipment and 'credit' not in current_shipment['charges']:
        shipments.append(current_shipment)

    # Generate final DataFrame...
    if not shipments:
         return pd.DataFrame()
    df = pd.DataFrame(shipments)
    if 'charges' in df.columns:
        charges_df = pd.json_normalize(df['charges'])
        final_df = pd.concat([df.drop(columns=['charges']).reset_index(drop=True),
                              charges_df.reset_index(drop=True)], axis=1)
    else:
        final_df = df
    return final_df



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
            df = extract_shipments(pdf_path) # Call your extraction function
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