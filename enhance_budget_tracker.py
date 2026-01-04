import pandas as pd
from datetime import datetime, timedelta
import os

# Create a sample enhanced expense tracker
def create_enhanced_expense_tracker():
    # Define the enhanced columns
    columns = [
        'Date',
        'Description',
        'Amount',
        'Expense Type',  # Subscription or One-Time
        'Billing Cycle',  # Monthly/Quarterly/Yearly/One-Time
        'Next Billing Date',
        'Subscription Status',  # Active/Cancelled/Paused
        'Category',
        'Payment Method',
        'Recurring',  # Y/N
        'Notes',
        'Receipt'  # For receipt file paths or links
    ]
    
    # Create sample data
    today = datetime.now().date()
    next_month = (today.replace(day=1) + timedelta(days=32)).replace(day=1)
    
    data = [
        # Sample subscription
        [
            today - timedelta(days=5),  # Date
            'Netflix Subscription',  # Description
            15.99,  # Amount
            'Subscription',  # Expense Type
            'Monthly',  # Billing Cycle
            next_month,  # Next Billing Date
            'Active',  # Subscription Status
            'Entertainment',  # Category
            'Credit Card',  # Payment Method
            'Y',  # Recurring
            'Premium Plan',  # Notes
            ''  # Receipt
        ],
        # Sample one-time purchase
        [
            today - timedelta(days=2),  # Date
            'Office Chair',  # Description
            199.99,  # Amount
            'One-Time',  # Expense Type
            'One-Time',  # Billing Cycle
            None,  # Next Billing Date
            'N/A',  # Subscription Status
            'Furniture',  # Category
            'Debit Card',  # Payment Method
            'N',  # Recurring
            'Ergonomic chair for home office',  # Notes
            'receipt123.jpg'  # Receipt
        ]
    ]
    
    # Create DataFrame
    df = pd.DataFrame(data, columns=columns)
    
    # Format dates
    df['Date'] = pd.to_datetime(df['Date']).dt.date
    df['Next Billing Date'] = pd.to_datetime(df['Next Billing Date']).dt.date
    
    return df

def create_enhanced_budget_template(output_file):
    # Create a Pandas Excel writer
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Create the enhanced expense tracker
        expense_df = create_enhanced_expense_tracker()
        expense_df.to_excel(writer, sheet_name='Expense Tracker', index=False)
        
        # Create a data validation sheet for dropdowns
        validation_data = {
            'Expense Type': ['Subscription', 'One-Time'],
            'Billing Cycle': ['Monthly', 'Quarterly', 'Yearly', 'One-Time'],
            'Subscription Status': ['Active', 'Cancelled', 'Paused', 'N/A'],
            'Category': [
                'Housing', 'Utilities', 'Groceries', 'Dining Out', 'Transportation',
                'Health', 'Insurance', 'Personal Care', 'Entertainment', 'Education',
                'Shopping', 'Gifts', 'Travel', 'Subscriptions', 'Investments', 'Other'
            ],
            'Payment Method': [
                'Cash', 'Credit Card', 'Debit Card', 'Bank Transfer', 'PayPal',
                'Mobile Payment', 'Cryptocurrency', 'Other'
            ],
            'Recurring': ['Y', 'N']
        }
        
        # Create validation sheet
        max_len = max(len(v) for v in validation_data.values())
        for key in validation_data:
            validation_data[key] += [''] * (max_len - len(validation_data[key]))
            
        validation_df = pd.DataFrame(validation_data)
        validation_df.to_excel(writer, sheet_name='Validation Lists', index=False)
        
        # Create a summary sheet
        summary_data = {
            'Metric': [
                'Total Monthly Subscriptions',
                'Total One-Time Expenses',
                'Upcoming Renewals (Next 30 days)',
                'Most Expensive Category',
                'Average Monthly Spend'
            ],
            'Value': [''] * 5,
            'Last Updated': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * 5
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Create a subscriptions overview sheet
        subscriptions_df = expense_df[expense_df['Expense Type'] == 'Subscription'].copy()
        if not subscriptions_df.empty:
            subscriptions_df.to_excel(writer, sheet_name='Subscriptions', index=False)

if __name__ == "__main__":
    output_file = 'Enhanced_Budget_Tracker.xlsx'
    create_enhanced_budget_template(output_file)
    print(f"Enhanced budget template created: {output_file}")
