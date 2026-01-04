#!/usr/bin/env python3

import PyPDF2
import re
import datetime

def test_new_parser(file_path):
    """Test the new Monzo PDF parsing logic"""
    print(f"=== TESTING NEW MONZO PDF PARSER ===")
    
    try:
        # Read PDF
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            
            # Extract text from all pages
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            
            print(f"Total extracted text length: {len(text)} characters")
            
            # Use regex to find all transaction patterns in the concatenated text
            # Pattern: DD/MM/YYYYDescriptionAmountBalanceAmount
            transaction_pattern = r'(\d{2}/\d{2}/\d{4})([A-Za-z0-9\s\.\-\(\)\/]+?)(-?\d+\.\d{2})(-?\d+\.\d{2})'
            
            matches = re.findall(transaction_pattern, text)
            print(f"Found {len(matches)} transaction matches using regex")
            
            all_transactions = []
            
            for i, match in enumerate(matches[:20]):  # Show first 20
                date_str, description, amount_str, balance_str = match
                
                try:
                    amount = float(amount_str)
                    balance = float(balance_str)
                    
                    # Clean up description
                    description = description.strip()
                    
                    print(f"{i+1}. {date_str} | {amount:>8.2f} | {description}")
                    
                    all_transactions.append({
                        'Date': date_str,
                        'Description': description,
                        'Amount': amount,
                        'Balance': balance
                    })
                    
                except (ValueError, IndexError) as e:
                    print(f"Error parsing transaction {i}: {match}, Error: {e}")
                    continue
            
            # If regex didn't work well, try alternative approach
            if len(all_transactions) < 10:
                print("\nRegex approach found few transactions, trying alternative parsing...")
                
                # Alternative: split by date pattern
                date_pattern = r'(\d{2}/\d{2}/\d{4})'
                parts = re.split(date_pattern, text)
                
                print(f"Split text into {len(parts)} parts")
                
                alt_transactions = []
                current_date = None
                for i in range(1, min(len(parts), 40)):  # Show first 20 transactions
                    if i + 1 < len(parts):
                        date_str = parts[i]
                        content = parts[i + 1]
                        
                        try:
                            current_date = datetime.datetime.strptime(date_str, '%d/%m/%Y').strftime('%Y-%m-%d')
                            
                            # Find all amounts in this section
                            amount_matches = re.findall(r'(-?\d+\.\d{2})', content)
                            
                            if len(amount_matches) >= 2:  # Need at least amount and balance
                                amount = float(amount_matches[0])
                                
                                # Extract description (text between date and first amount)
                                description = content[:content.find(amount_matches[0])].strip()
                                
                                if len(description) >= 2 and abs(amount) >= 0.01:
                                    print(f"Alt {len(alt_transactions)+1}. {date_str} | {amount:>8.2f} | {description}")
                                    
                                    alt_transactions.append({
                                        'Date': date_str,
                                        'Description': description,
                                        'Amount': amount
                                    })
                        
                        except (ValueError, IndexError):
                            continue
                
                print(f"\nAlternative method found {len(alt_transactions)} transactions")
            
            return all_transactions
            
    except Exception as e:
        print(f"Error: {str(e)}")
        return []

if __name__ == "__main__":
    # Test with your Monzo statement
    pdf_path = "/Users/anthonygathukia/Desktop/Me/Finance Folder's/Finance Budget Script/Test Site/Monzo_bank_statement_2025-11-01-2025-11-30_722.pdf"
    transactions = test_new_parser(pdf_path)
