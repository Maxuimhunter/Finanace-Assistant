#!/usr/bin/env python3

import PyPDF2
import re
import datetime

def debug_monzo_pdf(file_path):
    """Debug Monzo PDF parsing to see exactly what's extracted"""
    print(f"=== DEBUGGING MONZO PDF: {file_path} ===")
    
    try:
        # Read PDF
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            
            # Extract text from all pages
            for page_num, page in enumerate(pdf_reader.pages):
                page_text = page.extract_text()
                text += f"\n--- PAGE {page_num + 1} ---\n{page_text}\n"
            
            print(f"Total pages: {len(pdf_reader.pages)}")
            print(f"Total extracted text length: {len(text)} characters")
            
            # Show first 2000 characters of extracted text
            print("\n=== FIRST 2000 CHARACTERS OF EXTRACTED TEXT ===")
            print(text[:2000])
            print("=== END OF PREVIEW ===\n")
            
            # Parse line by line with detailed logging
            lines = text.split('\n')
            print(f"Total lines: {len(lines)}")
            
            all_transactions = []
            skipped_lines = []
            amount_lines = []
            
            for line_num, line in enumerate(lines):
                line = line.strip()
                
                if not line or len(line) < 5:
                    skipped_lines.append(f"Line {line_num}: Empty/short line")
                    continue
                
                # Check for skip patterns
                skip_reasons = []
                skip_patterns = [
                    'Personal Account statement', 'Total balance', 'Account number', 'BIC:', 'Sort code:',
                    'Date Description(GBP)', 'Amount(GBP)', 'Balance', 'Pot balance', 'Total outgoings',
                    'Total deposits', 'Monzo Bank Limited', 'Important information', 'Financial Services Register',
                    'Instant Access Savings Pot', 'Regular Pot provided by Monzo'
                ]
                
                for pattern in skip_patterns:
                    if pattern in line:
                        skip_reasons.append(pattern)
                        break
                
                if skip_reasons:
                    skipped_lines.append(f"Line {line_num}: Skipped due to: {skip_reasons[0]} | '{line[:50]}...'")
                    continue
                
                # Look for amounts
                amount_matches = re.findall(r'(-?\d+\.\d{2})', line)
                if amount_matches:
                    amount_lines.append(f"Line {line_num}: Found amounts {amount_matches} | '{line}'")
                    
                    for amount_str in amount_matches:
                        try:
                            amount = float(amount_str)
                            
                            if abs(amount) < 0.01 or abs(amount) > 10000:
                                skipped_lines.append(f"Line {line_num}: Amount {amount} out of range | '{line}'")
                                continue
                            
                            # Extract description
                            amount_pos = line.find(amount_str)
                            if amount_pos > 0:
                                description = line[:amount_pos].strip()
                            else:
                                description = line.replace(amount_str, '').strip()
                            
                            # Find date
                            date_match = re.search(r'(\d{2}/\d{2}/\d{4})', line)
                            if date_match:
                                date_str = date_match.group(1)
                            else:
                                # Look for date in previous lines
                                for j in range(max(0, line_num-5), line_num):
                                    prev_date_match = re.search(r'(\d{2}/\d{2}/\d{4})', lines[j])
                                    if prev_date_match:
                                        date_str = prev_date_match.group(1)
                                        break
                                else:
                                    date_str = "01/11/2025"
                            
                            all_transactions.append({
                                'line': line_num,
                                'date': date_str,
                                'description': description,
                                'amount': amount,
                                'original_line': line
                            })
                            
                        except ValueError:
                            skipped_lines.append(f"Line {line_num}: Invalid amount {amount_str} | '{line}'")
                else:
                    skipped_lines.append(f"Line {line_num}: No amount found | '{line}'")
            
            print(f"\n=== SUMMARY ===")
            print(f"Total lines processed: {len(lines)}")
            print(f"Lines with amounts: {len(amount_lines)}")
            print(f"Raw transactions found: {len(all_transactions)}")
            print(f"Lines skipped: {len(skipped_lines)}")
            
            print(f"\n=== LINES WITH AMOUNTS ===")
            for line_info in amount_lines[:20]:  # Show first 20
                print(line_info)
            
            print(f"\n=== RAW TRANSACTIONS FOUND ===")
            for i, trans in enumerate(all_transactions[:20]):  # Show first 20
                print(f"{i+1}. {trans['date']} | {trans['amount']:>8.2f} | {trans['description']}")
            
            print(f"\n=== SAMPLE SKIPPED LINES ===")
            for line_info in skipped_lines[:20]:  # Show first 20
                print(line_info)
            
            if len(all_transactions) > 20:
                print(f"\n... and {len(all_transactions) - 20} more transactions")
                print(f"... and {len(skipped_lines) - 20} more skipped lines")
            
            return all_transactions
            
    except Exception as e:
        print(f"Error: {str(e)}")
        return []

if __name__ == "__main__":
    # Test with your Monzo statement
    pdf_path = "/Users/anthonygathukia/Desktop/Me/Finance Folder's/Finance Budget Script/Test Site/Monzo_bank_statement_2025-11-01-2025-11-30_722.pdf"
    transactions = debug_monzo_pdf(pdf_path)
