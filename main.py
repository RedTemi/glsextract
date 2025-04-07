import pdfplumber
import re
import pandas as pd

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

# Usage
df = extract_shipments("DICOM_INVOICE.pdf")
df.to_csv("processed_shipments.csv", index=False)