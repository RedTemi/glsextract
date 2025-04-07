import pdfplumber
import re
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

def extract_shipments(pdf_path):
    shipments = []
    current_shipment = None
    
    # Structured charge patterns with uniform handling
    charge_patterns = [
        # Basic Charges
        (r'\bweight\b', 'weight_charge'),
        (r'carbon\s*surch?r[gq]?\.?', 'carbon surchrg.'),
        (r'\bfuel\b', 'fuel'),
        
        # Special Services
        (r'2nd\s*delivery', '2nd delivery'),
        (r'adrs\s*correction', 'adrs correction'),
        (r'ps:\s*max\s*limits', 'ps: max limits'),
        
        # Size/Weight Surcharges
        (r'non.?conveyable', 'non-conveyable'),
        (r'over\s*36\s*inches', 'over 36 inches'),
        (r'over\s*44\s*inches', 'over 44 inches'),
        (r'over\s*max\s*limits', 'over max limits'),
        (r'overweight\s*\(pc\)', 'overweight (pc)'),
        (r'overweight\s*\(sh\)', 'overweight (sh)'),
        
        # Taxes (with dynamic handling)
        (r'HST\s*(NB|NFL|NS|ON|PE)?', 
            lambda m: f'hst {m.group(1).lower()}' if m.group(1) else 'hst'),
        
        # Miscellaneous
        (r'\bcredit\b', 'credit'),
        (r'\bzone\b', 'zone'),
        (r'GST', 'GST'),
        (r'QST', 'QST')
    ]

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            for line in lines:
                # Waybill detection and processing
                waybill_match = re.match(
                    r'^([A-Z]\d{8})\s+'
                    r'(\d{2}\s\d{2}\s\d{4})\s+'
                    r'\d+\s+'
                    r'(\d+\.\d+)\s+'
                    r'(lb|kg)\s+'
                    r'(.*?)\s+'
                    r'base\s+'
                    r'(\d+\.\d+)', 
                    line
                )
                
                if waybill_match:
                    if current_shipment:
                        shipments.append(current_shipment)
                    
                    description = waybill_match.group(5).strip().lower()
                    if description == "adjustment":
                        current_shipment = None
                        continue
                        
                    current_shipment = {
                        'waybill': waybill_match.group(1),
                        'date': '-'.join(waybill_match.group(2).split()),
                        'weight': float(waybill_match.group(3)),
                        'unit': waybill_match.group(4),
                        'description': description,
                        'base': float(waybill_match.group(6)),
                        'charges': {},
                        'total': None
                    }
                    continue
                
                if current_shipment:
                    # Process charge lines
                    matches = re.findall(r'([A-Za-z][\w\s\-\(\):]+?)\s+(\d+\.\d+)', line)
                    for charge_raw, amount in matches:
                        charge_name = None
                        for pattern, handler in charge_patterns:
                            match = re.search(pattern, charge_raw, re.IGNORECASE)
                            if match:
                                charge_name = handler(match) if callable(handler) else handler
                                break
                        
                        if charge_name and charge_name != 'credit':
                            current_shipment['charges'][charge_name] = float(amount)
                    
                    total_match = re.search(r'(\d+\.\d+)\s*$', line)
                    if total_match:
                        current_shipment['total'] = float(total_match.group(1))
    
    if current_shipment and 'credit' not in current_shipment['charges']:
        shipments.append(current_shipment)

    # Generate final DataFrame
    df = pd.DataFrame(shipments)
    if not df.empty:
        charges_df = pd.json_normalize(df['charges'])
        final_df = pd.concat([df.drop(columns=['charges']), charges_df], axis=1)
        return final_df
    return pd.DataFrame()


def process_folder(input_folder, output_path, status_label):
    all_dfs = []
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        status_label.config(text="No PDF files found in selected folder")
        return

    status_label.config(text=f"Processing {len(pdf_files)} PDFs...")
    
    for i, pdf_file in enumerate(pdf_files):
        try:
            pdf_path = os.path.join(input_folder, pdf_file)
            df = extract_shipments(pdf_path)
            if not df.empty:
                all_dfs.append(df)
            status_label.config(text=f"Processed {i+1}/{len(pdf_files)} files")
        except Exception as e:
            print(f"Error processing {pdf_file}: {str(e)}")
    
    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        final_df.to_csv(output_path, index=False)
        status_label.config(text=f"Success! Combined CSV saved to {output_path}")
        messagebox.showinfo("Success", f"Combined CSV saved to:\n{output_path}")
    else:
        status_label.config(text="No valid data found in PDFs")
        messagebox.showwarning("No Data", "No shipment data found in PDF files")

class PDFApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Shipment Extractor")
        self.geometry("500x200")
        
        # Input Folder
        self.input_folder = tk.StringVar()
        ttk.Label(self, text="Input Folder:").pack(pady=5)
        ttk.Entry(self, textvariable=self.input_folder, width=50).pack()
        ttk.Button(self, text="Browse...", command=self.select_input_folder).pack()
        
        # Output File
        self.output_file = tk.StringVar()
        ttk.Label(self, text="Output CSV:").pack(pady=5)
        ttk.Entry(self, textvariable=self.output_file, width=50).pack()
        ttk.Button(self, text="Browse...", command=self.select_output_file).pack()
        
        # Status
        self.status_label = ttk.Label(self, text="")
        self.status_label.pack(pady=10)
        
        # Process Button
        ttk.Button(self, text="Process PDFs", command=self.start_processing).pack()
    
    def select_input_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_folder.set(folder)
            default_output = os.path.join(folder, "combined_shipments.csv")
            self.output_file.set(default_output)
    
    def select_output_file(self):
        file = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )
        if file:
            self.output_file.set(file)
    
    def start_processing(self):
        input_folder = self.input_folder.get()
        output_file = self.output_file.get()
        
        if not input_folder or not output_file:
            messagebox.showerror("Error", "Please select both input folder and output file")
            return
        
        Thread(target=process_folder, args=(
            input_folder,
            output_file,
            self.status_label
        ), daemon=True).start()

if __name__ == "__main__":
    app = PDFApp()
    app.mainloop()