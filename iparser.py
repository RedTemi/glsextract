import pdfplumber
import re
import pandas as pd

def extract_shipments(pdf_path):
    shipments = []
    current_shipment = None
    charge_patterns = {
        # Basic charges
        r'\bbase\b': 'base',
        r'\bweight\b': 'weight',
        r'carbon\s*surch?r[gq]?\.?': 'carbon surchrg.',
        r'fuel': 'fuel',
        
        # Special handling for multi-word charges
        r'non.?conveyable': 'non-conveyable',
        r'over\s*36\s*inches': 'over 36 inches',
        r'over\s*44\s*inches': 'over 44 inches',
        r'overweight\s*\(pc\)': 'overweight (pc)',
        r'overweight\s*\(sh\)': 'overweight (sh)',
        
        # Taxes
        r'GST': 'GST',
        r'QST': 'QST',
        r'HST\s*NB': 'hst nb',
        r'HST\s*NFL': 'hst nfl', 
        r'HST\s*NS': 'hst ns',
        r'HST\s*ON': 'hst on',
        r'HST\s*PE': 'hst pe',
        r'HST\b': 'hst'
    }

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            for line in lines:
                # Start new shipment when waybill is found
                if re.match(r'^[NP]\d+', line):
                    if current_shipment: 
                        shipments.append(current_shipment)
                        
                    parts = re.split(r'\s{2,}', line)
                    current_shipment = {
                        'waybill': re.search(r'([NP]\d+)', line).group(1),
                        'date': '-'.join(re.findall(r'\d{2}\s\d{2}\s\d{4}', line)[0].split()),
                        'weight': float(re.search(r'(\d+\.\d+)\s*lb', line).group(1)),
                        'unit': 'lb',
                        'description': ' '.join(parts[5:]) if len(parts) > 5 else None,
                        'charges': {},
                        'total': None
                    }
                    continue
                
                # Process charge lines
                if current_shipment:
                    # Match charges with amounts (e.g., "carbon surchrg. 0.90")
                    for pattern, name in charge_patterns.items():
                        match = re.search(rf'({pattern})\s+(\d+\.\d+)', line, re.IGNORECASE)
                        if match:
                            current_shipment['charges'][name] = float(match.group(2))
                    
                    # Capture total at line end (e.g., "QST 4.40 50.69")
                    total_match = re.search(r'(\d+\.\d+)\s*$', line)
                    if total_match:
                        current_shipment['total'] = float(total_match.group(1))
    
    # Convert to DataFrame
    df = pd.DataFrame(shipments)
    charges_df = pd.json_normalize(df['charges'])
    final_df = pd.concat([df.drop(['charges'], axis=1), charges_df], axis=1)
    return final_df

# Usage
df = extract_shipments("DICOM_INVOICE.pdf")
df.to_csv("processed_shipments.csv", index=False)