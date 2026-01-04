import generator

# Test the improved PDF parser
with open('Monzo_bank_statement_2025-11-01-2025-11-30_722.pdf', 'rb') as f:
    file_content = f.read()
    
    df = generator.parse_monzo_statement(file_content)
    
    if df is not None:
        print(f"✅ Successfully parsed {len(df)} transactions")
        print(f"Date range: {df['Date'].min()} to {df['Date'].max()}")
        print("\nAll transactions:")
        for i, row in df.iterrows():
            desc = row['Description']
            if len(desc) > 30:
                desc = desc[:30] + "..."
            else:
                desc = desc
            print(f"{i+1}: {row['Date']} | {desc:30} | ${row['Amount']:>8.2f} | {row['Category']}")
        
        print("\nSummary by category:")
        category_summary = df.groupby('Category')['Amount'].sum()
        for cat, amount in category_summary.items():
            print(f"{cat}: ${abs(amount):.2f}")
    else:
        print("❌ Failed to parse PDF")
