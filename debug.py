import pdfplumber
import re

def debug_pdf(pdf_path):
    with open("debug.log", "w") as f:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                f.write(f"\n=== PAGE {page_num + 1} ===\n")
                text = page.extract_text()
                f.write(text + "\n")
                
                # Look for table headers
                lines = text.split('\n')
                header_idx = None
                for idx, line in enumerate(lines):
                    if "WAYBILL" in line and "TOTAL $" in line:
                        header_idx = idx
                        f.write(f"\nHEADER FOUND AT LINE {idx}: {line}\n")
                        break
                
                if header_idx is None:
                    f.write("\nNO HEADER DETECTED ON THIS PAGE\n")
                    continue
                
                # Process potential data rows
                for line_num, line in enumerate(lines[header_idx+1:], start=header_idx+1):
                    f.write(f"\nROW {line_num}: {line}\n")
                    
                    # Column splitting test
                    cols = re.split(r'\s*\|\s*', line.strip())
                    f.write(f"Split columns: {cols}\n")
                    
                    # Waybill detection test
                    waybill_match = re.search(r'[NP]\d+', line)
                    f.write(f"Waybill match: {waybill_match.group() if waybill_match else 'FAILED'}\n")

# Usage
debug_pdf("DICOM_INVOICE.pdf")