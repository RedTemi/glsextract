import pdfplumber
import re
import pandas as pd

def extract_shipments(pdf_path):
    shipments = []
    current_shipment = None
    
    charge_patterns = {
        # Basic Charges
        r'\bbase\b': 'base',
        r'\bweight\b': 'weight_charge',
        r'carbon\s*surch?r[gq]?\.?': 'carbon surchrg.',
        r'\bfuel\b': 'fuel',
        
        # Special Services
        r'2nd\s*delivery': '2nd delivery',
        r'adrs\s*correction': 'adrs correction',
        r'ps:\s*max\s*limits': 'ps: max limits',
        
        # Size/Weight Surcharges
        r'non.?conveyable': 'non-conveyable',
        r'over\s*36\s*inches': 'over 36 inches',
        r'over\s*44\s*inches': 'over 44 inches', 
        r'over\s*max\s*limits': 'over max limits',
        r'overweight\s*\(pc\)': 'overweight (pc)',
        r'overweight\s*\(sh\)': 'overweight (sh)',
        
        # Taxes
        r'GST': 'GST',
        r'QST': 'QST',
        r'HST\s*(NB|NFL|NS|ON|PE)?': lambda m: f'hst {m.group(1).lower()}' if m.group(1) else 'hst',
        
        # Miscellaneous
        r'\bcredit\b': 'credit',
        r'\bzone\b': 'zone'
    }

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            for line in lines:
                # Detect Waybill line with base charge (strict format)
                waybill_match = re.match(
                    r'^([A-Z]\d{8})\s+'          # Waybill ID
                    r'(\d{2}\s\d{2}\s\d{4})\s+'  # Date
                    r'\d+\s+'                    # Pieces
                    r'(\d+\.\d+)\s+'             # Weight
                    r'(lb|kg)\s+'                # Unit
                    r'(.*?)\s+'                  # Description
                    r'base\s+'                   # Base charge label
                    r'(\d+\.\d+)',               # Base charge amount
                    line
                )
                
                if waybill_match:
                    if current_shipment:
                        shipments.append(current_shipment)
                    
                    current_shipment = {
                        'waybill': waybill_match.group(1),
                        'date': '-'.join(waybill_match.group(2).split()),
                        'weight': float(waybill_match.group(3)),
                        'unit': waybill_match.group(4),
                        'description': waybill_match.group(5).strip(),
                        'base': float(waybill_match.group(6)),
                        'charges': {},
                        'total': None
                    }
                    continue
                
                # Process charge lines for current shipment
                if current_shipment:
                    # Extract all charge-amount pairs
                    matches = re.findall(
                        r'([A-Za-z][\w\s\-\(\):]+?)\s+(\d+\.\d+)', 
                        line
                    )
                    
                    for charge_raw, amount in matches:
                        # Standardize charge names
                        charge_name = None
                        for pattern, name in charge_patterns.items():
                            if callable(name):
                                match = re.search(pattern, charge_raw, re.IGNORECASE)
                                if match:
                                    charge_name = name(match)
                                    break
                            elif re.search(pattern, charge_raw, re.IGNORECASE):
                                charge_name = name
                                break
                        
                        if charge_name:
                            current_shipment['charges'][charge_name] = float(amount)
                    
                    # Capture total at line end
                    total_match = re.search(r'(\d+\.\d+)\s*$', line)
                    if total_match:
                        current_shipment['total'] = float(total_match.group(1))
    
    # Add final shipment
    if current_shipment:
        shipments.append(current_shipment)

    # Create comprehensive DataFrame
    all_columns = [
        'waybill', 'date', 'weight', 'unit', 'description', 'base', 'total',
        '2nd delivery', 'adrs correction', 'carbon surchrg.', 'credit', 'fuel',
        'non-conveyable', 'over 36 inches', 'over 44 inches', 'over max limits',
        'overweight (pc)', 'overweight (sh)', 'ps: max limits', 'zone',
        'GST', 'QST', 'hst', 'hst nb', 'hst nfl', 'hst ns', 'hst on', 'hst pe'
    ]
    
    df = pd.DataFrame(shipments)
    charges_df = pd.json_normalize(df['charges'])
    final_df = pd.concat([df.drop(columns=['charges']), charges_df], axis=1)
    return final_df.reindex(columns=all_columns).fillna(0)

# Usage
df = extract_shipments("DICOM_INVOICE.pdf")
df.to_csv("processed_shipments.csv", index=False)