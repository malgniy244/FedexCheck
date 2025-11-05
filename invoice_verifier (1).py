import pandas as pd
import pdfplumber
import re
from datetime import datetime

class ComprehensiveInvoiceVerifier:
    def __init__(self, excel_path, pdf_path):
        self.excel_path = excel_path
        self.pdf_path = pdf_path
        self.verification_results = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'contact_name': {},
            'purpose_of_shipment': {},
            'line_items': {'summary': {}, 'details': []}
        }
        
    def extract_excel_data(self):
        """Extract data from YOUR Excel invoice with all 7 fields"""
        print("📋 Reading YOUR Excel invoice...")
        
        df = pd.read_excel(self.excel_path, header=None)
        
        excel_data = {
            'contact_name': None,
            'purpose_of_shipment': None,
            'line_items': []
        }
        
        # Find Contact Name (row 7, col 1)
        if len(df) > 7 and len(df.columns) > 1:
            contact_val = df.iloc[7, 1]
            if pd.notna(contact_val):
                excel_data['contact_name'] = str(contact_val).strip()
        
        # Extract line items - find the header row first
        line_item_start = None
        for idx, row in df.iterrows():
            if pd.notna(row[0]) and str(row[0]).strip() == 'Rank':
                line_item_start = idx + 1
                break
        
        if line_item_start:
            for idx in range(line_item_start, len(df)):
                row = df.iloc[idx]
                
                # Check if this is the end
                if pd.notna(row[0]) and 'Total' in str(row[0]):
                    break
                
                # Extract all fields using correct column indices based on your Excel structure
                # Col 0: Rank, Col 1: Quantity, Col 2: HTS Code, Col 3: Country Full Name
                # Col 4: Country Code, Col 5: Year, Col 6: Description, Col 7: Material Type
                # Col 8: Weight (kgs), Col 9: Weight (gms), Col 10: Weight (Lbs), Col 11: Value (US $)
                
                rank = row[0] if pd.notna(row[0]) else None
                quantity = row[1] if pd.notna(row[1]) else None
                hs_code = row[2] if pd.notna(row[2]) else None
                country_full = row[3] if pd.notna(row[3]) else None
                country_code = row[4] if pd.notna(row[4]) else None
                year = row[5] if pd.notna(row[5]) else None
                description = row[6] if pd.notna(row[6]) else None
                net_weight_kg = row[8] if len(row) > 8 and pd.notna(row[8]) else None  # Weight in KG
                total_value = row[11] if len(row) > 11 and pd.notna(row[11]) else None  # Total Value
                
                # Calculate unit value
                unit_value = None
                if total_value and quantity and quantity > 0:
                    try:
                        unit_value = float(total_value) / float(quantity)
                    except:
                        pass
                
                if rank and quantity and description:
                    excel_data['line_items'].append({
                        'rank': int(rank) if str(rank).replace('.','').isdigit() else rank,
                        'quantity': float(quantity),
                        'net_weight': float(net_weight_kg) if net_weight_kg else None,
                        'description': str(description).strip(),
                        'hs_code': str(hs_code).replace('.', '') if hs_code else None,  # Remove dot from 9705.31
                        'country_code': str(country_code).strip() if country_code else None,
                        'year': int(year) if year else None,
                        'unit_value': round(unit_value, 6) if unit_value else None,
                        'total_value': float(total_value) if total_value else None
                    })
        
        print(f"  ✓ Contact: '{excel_data['contact_name']}'")
        print(f"  ✓ Found {len(excel_data['line_items'])} line items")
        
        return excel_data
    
    def extract_pdf_data(self):
        """Extract data from FedEx PDF invoice with all 7 fields"""
        print("📄 Reading FedEx PDF invoice...")
        
        pdf_data = {
            'contact_name': None,
            'purpose_of_shipment': None,
            'line_items': []
        }
        
        with pdfplumber.open(self.pdf_path) as pdf:
            all_text = ""
            
            for page in pdf.pages:
                text = page.extract_text()
                all_text += text + "\n"
            
            lines = all_text.split('\n')
            
            # Extract line items with ALL fields
            # Pattern: QTY WEIGHT PCS Description... HS_CODE COUNTRY UNIT_VALUE TOTAL_VALUE
            # Example: "346.00 3.46 PCS Collectors pieces... 970531 GB 5.000000 1,730.00"
            
            for i, line in enumerate(lines):
                # Look for lines starting with number + number + PCS + Collectors
                if re.search(r'^\d+\.?\d*\s+\d+\.?\d*\s+PCS\s+Collectors', line):
                    
                    # Extract all numeric values from the line using comprehensive regex
                    # Pattern: QTY WEIGHT PCS ... HS_CODE COUNTRY UNIT_VAL TOTAL_VAL
                    pattern = r'^(\d+\.?\d*)\s+(\d+\.?\d*)\s+PCS\s+(.+?)\s+(\d{6})\s+([A-Z]{2})\s+([\d.]+)\s+([\d,]+\.?\d*)$'
                    match = re.search(pattern, line)
                    
                    if match:
                        quantity = float(match.group(1))
                        net_weight = float(match.group(2))
                        desc_start = match.group(3)
                        hs_code = match.group(4)
                        country_code = match.group(5)
                        unit_value = float(match.group(6))
                        total_value_str = match.group(7).replace(',', '')
                        total_value = float(total_value_str)
                    else:
                        # Fallback: extract piece by piece
                        parts = line.split()
                        
                        quantity = float(parts[0]) if len(parts) > 0 else None
                        net_weight = float(parts[1]) if len(parts) > 1 else None
                        
                        # Find HS code (970531)
                        hs_match = re.search(r'970531', line)
                        hs_code = '970531' if hs_match else None
                        
                        # Find country code (2 uppercase letters after HS code)
                        country_match = re.search(r'970531\s+([A-Z]{2})', line)
                        country_code = country_match.group(1) if country_match else None
                        
                        # Find unit value and total value (numbers after country code)
                        values_match = re.search(r'970531\s+[A-Z]{2}\s+([\d.]+)\s+([\d,]+\.?\d*)', line)
                        unit_value = float(values_match.group(1)) if values_match else None
                        total_value_str = values_match.group(2).replace(',', '') if values_match else None
                        total_value = float(total_value_str) if total_value_str else None
                    
                    # Get full description from current and next lines
                    desc_parts = []
                    year = None
                    
                    # Start with partial description from current line (before HS code)
                    desc_before_hs = re.sub(r'\d+\.?\d*\s+\d+\.?\d*\s+PCS\s+', '', line)
                    desc_before_hs = re.sub(r'\s*970531.*$', '', desc_before_hs)
                    if desc_before_hs.strip():
                        desc_parts.append(desc_before_hs.strip())
                    
                    # Look forward for continuation lines
                    for j in range(i + 1, min(i + 6, len(lines))):
                        next_line = lines[j].strip()
                        
                        # Stop at next item
                        if re.match(r'^\d+\.?\d*\s+\d+\.?\d*\s+PCS', next_line):
                            break
                        
                        # Stop at separators
                        if not next_line or 'Total' in next_line or '=' in next_line:
                            break
                        
                        desc_parts.append(next_line)
                        
                        # Check for year at end
                        year_match = re.search(r'(Note|Coin)\s+(19\d{2}|20\d{2})$', next_line)
                        if year_match:
                            year = int(year_match.group(2))
                            break
                    
                    # Combine description
                    full_description = ' '.join(desc_parts).strip()
                    
                    # Validate we have the key fields
                    if quantity and net_weight and hs_code and country_code:
                        pdf_data['line_items'].append({
                            'quantity': quantity,
                            'net_weight': net_weight,
                            'description': full_description,
                            'hs_code': hs_code,
                            'country_code': country_code,
                            'year': year,
                            'unit_value': unit_value,
                            'total_value': total_value
                        })
            
            # Extract contact name
            if 'SB-SHIPPING - PRN' in all_text:
                pattern = r'(SB-SHIPPING - PRN \d+)'
                match = re.search(pattern, all_text)
                if match:
                    pdf_data['contact_name'] = match.group(1).strip()
            
            if not pdf_data['contact_name'] and 'HK SB-SHIPPING' in all_text:
                pdf_data['contact_name'] = 'HK SB-SHIPPING'
            
            # Extract purpose
            if 'REPAIR_AND_RETURN' in all_text and 'SB-SHIPPING' in all_text:
                pdf_data['purpose_of_shipment'] = 'SB-SHIPPING - REPAIR_AND_RETURN'
        
        print(f"  ✓ Contact: '{pdf_data['contact_name']}'")
        print(f"  ✓ Purpose: '{pdf_data['purpose_of_shipment']}'")
        print(f"  ✓ Found {len(pdf_data['line_items'])} line items")
        
        return pdf_data
    
    def normalize_description(self, desc):
        """Normalize description for comparison"""
        if not desc:
            return ""
        desc_str = str(desc).strip().lower()
        desc_str = re.sub(r'\s+', ' ', desc_str)
        return desc_str
    
    def compare_values(self, val1, val2, tolerance=0.01):
        """Compare two numeric values with tolerance"""
        if val1 is None and val2 is None:
            return True
        if val1 is None or val2 is None:
            return False
        try:
            return abs(float(val1) - float(val2)) < tolerance
        except:
            return str(val1) == str(val2)
    
    def compare_line_items(self, excel_items, pdf_items):
        """Compare all 7 fields for each line item"""
        
        results = {
            'total_excel': len(excel_items),
            'total_pdf': len(pdf_items),
            'perfect_matches': 0,
            'mismatches': [],
            'unmatched_excel': [],
            'unmatched_pdf': []
        }
        
        matched_pdf_indices = set()
        matched_excel_indices = set()
        
        # Try to match each Excel item with PDF items
        for excel_idx, excel_item in enumerate(excel_items):
            best_match = None
            best_match_score = 0
            best_match_idx = None
            
            for pdf_idx, pdf_item in enumerate(pdf_items):
                if pdf_idx in matched_pdf_indices:
                    continue
                
                # Calculate match score
                score = 0
                differences = {}
                
                # Check each field
                qty_match = self.compare_values(excel_item['quantity'], pdf_item['quantity'])
                weight_match = self.compare_values(excel_item['net_weight'], pdf_item['net_weight'])
                hs_match = str(excel_item.get('hs_code', '970531')) == str(pdf_item.get('hs_code', '970531'))
                country_match = str(excel_item['country_code']) == str(pdf_item['country_code'])
                year_match = excel_item.get('year') == pdf_item.get('year')
                unit_val_match = self.compare_values(excel_item.get('unit_value'), pdf_item.get('unit_value'))
                total_val_match = self.compare_values(excel_item.get('total_value'), pdf_item.get('total_value'))
                
                # Description match (normalized)
                excel_desc_norm = self.normalize_description(excel_item['description'])
                pdf_desc_norm = self.normalize_description(pdf_item['description'])
                desc_match = excel_desc_norm in pdf_desc_norm or pdf_desc_norm in excel_desc_norm
                
                # Calculate score (out of 8)
                if qty_match: score += 1
                if weight_match: score += 1
                if desc_match: score += 1
                if hs_match: score += 1
                if country_match: score += 1
                if year_match: score += 1
                if unit_val_match: score += 1
                if total_val_match: score += 1
                
                differences = {
                    'quantity': qty_match,
                    'net_weight': weight_match,
                    'description': desc_match,
                    'hs_code': hs_match,
                    'country_code': country_match,
                    'year': year_match,
                    'unit_value': unit_val_match,
                    'total_value': total_val_match
                }
                
                # Perfect match = score 8
                if score == 8:
                    results['perfect_matches'] += 1
                    matched_pdf_indices.add(pdf_idx)
                    matched_excel_indices.add(excel_idx)
                    break
                
                # Keep track of best partial match
                if score > best_match_score and score >= 4:  # At least half fields match
                    best_match_score = score
                    best_match = pdf_item
                    best_match_idx = pdf_idx
                    best_differences = differences
            
            # If no perfect match but found partial match
            if excel_idx not in matched_excel_indices and best_match:
                results['mismatches'].append({
                    'excel_item': excel_item,
                    'pdf_item': best_match,
                    'differences': best_differences,
                    'match_score': f"{best_match_score}/8"
                })
                matched_pdf_indices.add(best_match_idx)
                matched_excel_indices.add(excel_idx)
        
        # Find completely unmatched items
        for idx, item in enumerate(excel_items):
            if idx not in matched_excel_indices:
                results['unmatched_excel'].append(item)
        
        for idx, item in enumerate(pdf_items):
            if idx not in matched_pdf_indices:
                results['unmatched_pdf'].append(item)
        
        return results
    
    def verify(self):
        """Main verification method"""
        print("\n🔍 COMPREHENSIVE INVOICE VERIFICATION")
        print("=" * 80)
        print("Checking: Quantity, Weight, Description, HS Code, Country, Year, Unit Value, Total Value")
        print("=" * 80 + "\n")
        
        # Extract data
        excel_data = self.extract_excel_data()
        pdf_data = self.extract_pdf_data()
        
        print("\n" + "=" * 80)
        print("VERIFICATION CHECKS")
        print("=" * 80 + "\n")
        
        # Check 1: Contact Name
        print("📌 CHECK 1: Contact Name")
        print(f"   Required: 'SB-SHIPPING - PRN 5789187'")
        print(f"   Excel:    '{excel_data['contact_name']}'")
        print(f"   PDF:      '{pdf_data['contact_name']}'")
        
        contact_pass = (pdf_data['contact_name'] == 'SB-SHIPPING - PRN 5789187')
        
        self.verification_results['contact_name'] = {
            'required': 'SB-SHIPPING - PRN 5789187',
            'excel_value': excel_data['contact_name'],
            'pdf_value': pdf_data['contact_name'],
            'status': '✓ PASS' if contact_pass else '✗ FAIL'
        }
        
        print(f"   Result: {self.verification_results['contact_name']['status']}\n")
        
        # Check 2: Purpose
        print("📌 CHECK 2: Purpose of Shipment")
        print(f"   Required: 'SB-SHIPPING - REPAIR_AND_RETURN'")
        print(f"   PDF:      '{pdf_data['purpose_of_shipment']}'")
        
        purpose_pass = (pdf_data['purpose_of_shipment'] == 'SB-SHIPPING - REPAIR_AND_RETURN')
        
        self.verification_results['purpose_of_shipment'] = {
            'required': 'SB-SHIPPING - REPAIR_AND_RETURN',
            'pdf_value': pdf_data['purpose_of_shipment'],
            'status': '✓ PASS' if purpose_pass else '✗ FAIL'
        }
        
        print(f"   Result: {self.verification_results['purpose_of_shipment']['status']}\n")
        
        # Check 3: Line Items (ALL 7 FIELDS)
        print("📌 CHECK 3: Declaration Content - ALL 7 FIELDS")
        print(f"   Your Excel:  {len(excel_data['line_items'])} items")
        print(f"   FedEx PDF:   {len(pdf_data['line_items'])} items")
        
        comparison = self.compare_line_items(excel_data['line_items'], pdf_data['line_items'])
        
        self.verification_results['line_items'] = {
            'summary': comparison,
            'excel_items': excel_data['line_items'],
            'pdf_items': pdf_data['line_items']
        }
        
        print(f"   Perfect Matches:  {comparison['perfect_matches']}")
        print(f"   Partial Matches:  {len(comparison['mismatches'])}")
        print(f"   Excel Only:       {len(comparison['unmatched_excel'])}")
        print(f"   PDF Only:         {len(comparison['unmatched_pdf'])}")
        
        all_match = (comparison['perfect_matches'] == len(excel_data['line_items']) and
                     len(comparison['mismatches']) == 0 and
                     len(comparison['unmatched_excel']) == 0 and
                     len(comparison['unmatched_pdf']) == 0)
        
        print(f"   Result: {'✓ ALL FIELDS MATCH PERFECTLY' if all_match else '✗ DISCREPANCIES FOUND'}\n")
        
        return self.verification_results
    
    def generate_report(self, output_path):
        """Generate detailed verification report"""
        results = self.verification_results
        lines = []
        
        lines.append("=" * 100)
        lines.append("COMPREHENSIVE INVOICE VERIFICATION REPORT")
        lines.append("Checking: Quantity | Net Weight | Description | HS Code | Country | Year | Unit Value | Total Value")
        lines.append("=" * 100)
        lines.append(f"Generated: {results['timestamp']}\n")
        
        # Contact Name
        lines.append("-" * 100)
        lines.append("1. CONTACT NAME VERIFICATION")
        lines.append("-" * 100)
        cn = results['contact_name']
        lines.append(f"Required: {cn['required']}")
        lines.append(f"Your Excel: {cn['excel_value']}")
        lines.append(f"FedEx PDF:  {cn['pdf_value']}")
        lines.append(f"Result: {cn['status']}\n")
        
        # Purpose
        lines.append("-" * 100)
        lines.append("2. PURPOSE OF SHIPMENT VERIFICATION")
        lines.append("-" * 100)
        pos = results['purpose_of_shipment']
        lines.append(f"Required: {pos['required']}")
        lines.append(f"FedEx PDF: {pos['pdf_value']}")
        lines.append(f"Result: {pos['status']}\n")
        
        # Line Items
        lines.append("-" * 100)
        lines.append("3. LINE ITEMS VERIFICATION - ALL 7 FIELDS")
        lines.append("-" * 100)
        summary = results['line_items']['summary']
        lines.append(f"Total in Excel:      {summary['total_excel']}")
        lines.append(f"Total in PDF:        {summary['total_pdf']}")
        lines.append(f"Perfect Matches:     {summary['perfect_matches']}")
        lines.append(f"Partial Matches:     {len(summary['mismatches'])}")
        lines.append(f"Excel Only:          {len(summary['unmatched_excel'])}")
        lines.append(f"PDF Only:            {len(summary['unmatched_pdf'])}\n")
        
        # Show mismatches in detail
        if summary['mismatches']:
            lines.append("=" * 100)
            lines.append("ITEMS WITH DISCREPANCIES (PARTIAL MATCHES)")
            lines.append("=" * 100 + "\n")
            
            for idx, mismatch in enumerate(summary['mismatches'], 1):
                excel_item = mismatch['excel_item']
                pdf_item = mismatch['pdf_item']
                diffs = mismatch['differences']
                
                lines.append(f"Discrepancy #{idx} - Match Score: {mismatch['match_score']}")
                lines.append("-" * 100)
                lines.append("YOUR EXCEL:")
                lines.append(f"  Quantity:     {excel_item.get('quantity')}")
                lines.append(f"  Net Weight:   {excel_item.get('net_weight')} kg")
                lines.append(f"  Description:  {excel_item.get('description')[:80]}...")
                lines.append(f"  HS Code:      {excel_item.get('hs_code')}")
                lines.append(f"  Country:      {excel_item.get('country_code')}")
                lines.append(f"  Year:         {excel_item.get('year')}")
                lines.append(f"  Unit Value:   ${excel_item.get('unit_value')}")
                lines.append(f"  Total Value:  ${excel_item.get('total_value')}")
                
                lines.append("\nFEDEX PDF:")
                lines.append(f"  Quantity:     {pdf_item.get('quantity')}")
                lines.append(f"  Net Weight:   {pdf_item.get('net_weight')} kg")
                lines.append(f"  Description:  {pdf_item.get('description')[:80]}...")
                lines.append(f"  HS Code:      {pdf_item.get('hs_code')}")
                lines.append(f"  Country:      {pdf_item.get('country_code')}")
                lines.append(f"  Year:         {pdf_item.get('year')}")
                lines.append(f"  Unit Value:   ${pdf_item.get('unit_value')}")
                lines.append(f"  Total Value:  ${pdf_item.get('total_value')}")
                
                lines.append("\nFIELD-BY-FIELD COMPARISON:")
                lines.append(f"  Quantity:     {'✓ Match' if diffs['quantity'] else '✗ DIFFERENT'}")
                lines.append(f"  Net Weight:   {'✓ Match' if diffs['net_weight'] else '✗ DIFFERENT'}")
                lines.append(f"  Description:  {'✓ Match' if diffs['description'] else '✗ DIFFERENT'}")
                lines.append(f"  HS Code:      {'✓ Match' if diffs['hs_code'] else '✗ DIFFERENT (Should all be 970531)'}")
                lines.append(f"  Country:      {'✓ Match' if diffs['country_code'] else '✗ DIFFERENT'}")
                lines.append(f"  Year:         {'✓ Match' if diffs['year'] else '✗ DIFFERENT'}")
                lines.append(f"  Unit Value:   {'✓ Match' if diffs['unit_value'] else '✗ DIFFERENT'}")
                lines.append(f"  Total Value:  {'✓ Match' if diffs['total_value'] else '✗ DIFFERENT'}")
                lines.append("\n")
        
        # Show unmatched Excel items
        if summary['unmatched_excel']:
            lines.append("=" * 100)
            lines.append(f"ITEMS ONLY IN YOUR EXCEL ({len(summary['unmatched_excel'])} items - NOT FOUND IN FEDEX PDF)")
            lines.append("=" * 100 + "\n")
            
            for idx, item in enumerate(summary['unmatched_excel'], 1):
                lines.append(f"Excel Item #{idx}:")
                lines.append(f"  Quantity:     {item.get('quantity')}")
                lines.append(f"  Net Weight:   {item.get('net_weight')} kg")
                lines.append(f"  Description:  {item.get('description')[:80]}...")
                lines.append(f"  HS Code:      {item.get('hs_code')}")
                lines.append(f"  Country:      {item.get('country_code')}")
                lines.append(f"  Year:         {item.get('year')}")
                lines.append(f"  Unit Value:   ${item.get('unit_value')}")
                lines.append(f"  Total Value:  ${item.get('total_value')}")
                lines.append("")
        
        # Show unmatched PDF items
        if summary['unmatched_pdf']:
            lines.append("=" * 100)
            lines.append(f"ITEMS ONLY IN FEDEX PDF ({len(summary['unmatched_pdf'])} items - NOT FOUND IN YOUR EXCEL)")
            lines.append("=" * 100 + "\n")
            
            for idx, item in enumerate(summary['unmatched_pdf'], 1):
                lines.append(f"FedEx PDF Item #{idx}:")
                lines.append(f"  Quantity:     {item.get('quantity')}")
                lines.append(f"  Net Weight:   {item.get('net_weight')} kg")
                lines.append(f"  Description:  {item.get('description')[:80]}...")
                lines.append(f"  HS Code:      {item.get('hs_code')}")
                lines.append(f"  Country:      {item.get('country_code')}")
                lines.append(f"  Year:         {item.get('year')}")
                lines.append(f"  Unit Value:   ${item.get('unit_value')}")
                lines.append(f"  Total Value:  ${item.get('total_value')}")
                lines.append("")
        
        # Final Summary
        lines.append("=" * 100)
        lines.append("FINAL VERIFICATION SUMMARY")
        lines.append("=" * 100)
        
        contact_ok = results['contact_name']['status'] == '✓ PASS'
        purpose_ok = results['purpose_of_shipment']['status'] == '✓ PASS'
        items_ok = (summary['perfect_matches'] == summary['total_excel'] and
                    len(summary['mismatches']) == 0 and
                    len(summary['unmatched_excel']) == 0 and
                    len(summary['unmatched_pdf']) == 0)
        
        lines.append(f"Contact Name:        {results['contact_name']['status']}")
        lines.append(f"Purpose of Shipment: {results['purpose_of_shipment']['status']}")
        lines.append(f"All Line Items:      {'✓ PASS - All 7 fields match perfectly' if items_ok else '✗ FAIL - See discrepancies above'}")
        lines.append(f"Perfect Matches:     {summary['perfect_matches']}/{summary['total_excel']}")
        lines.append("")
        
        if contact_ok and purpose_ok and items_ok:
            lines.append("╔═══════════════════════════════════════════════════════════════════════════════════════╗")
            lines.append("║  ✓✓✓ ALL VERIFICATION CHECKS PASSED ✓✓✓                                              ║")
            lines.append("║  Your invoice matches FedEx invoice perfectly on ALL 7 FIELDS!                       ║")
            lines.append("║  ✓ Quantity ✓ Weight ✓ Description ✓ HS Code ✓ Country ✓ Year ✓ Unit Value ✓ Total ║")
            lines.append("╚═══════════════════════════════════════════════════════════════════════════════════════╝")
        else:
            lines.append("╔═══════════════════════════════════════════════════════════════════════════════════════╗")
            lines.append("║  ✗✗✗ VERIFICATION FAILED - DISCREPANCIES FOUND ✗✗✗                                   ║")
            lines.append("║  Please review the detailed discrepancies above.                                     ║")
            lines.append("║  Check: Quantity, Weight, Description, HS Code, Country, Year, Unit Value, Total     ║")
            lines.append("╚═══════════════════════════════════════════════════════════════════════════════════════╝")
        
        lines.append("")
        
        report_text = "\n".join(lines)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(report_text)
        
        print("\n" + report_text)
        
        return report_text


if __name__ == "__main__":
    import sys
    
    # For testing
    excel_file = "/home/claude/invoice.xlsx"
    pdf_file = "/mnt/user-data/uploads/testing1.pdf"
    
    verifier = ComprehensiveInvoiceVerifier(excel_file, pdf_file)
    verifier.verify()
    
    report_path = "/home/claude/comprehensive_verification_report.txt"
    verifier.generate_report(report_path)
    
    print(f"\n✅ Comprehensive verification complete!")
    print(f"📄 Full report: {report_path}")
