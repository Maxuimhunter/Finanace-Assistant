def create_excel_template(month=None, sections=None):
    """
    Create a simple Excel template for a specific month with specified sections.
    This version is simplified to avoid file corruption issues.
    """
    # Define constants
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
             'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    # If no month specified, use current month
    if month is None:
        current_month = datetime.date.today().month - 1  # 0-based index
        month = months[current_month]
    
    # Create a new workbook
    wb = openpyxl.Workbook()
    # Remove default sheet
    while len(wb.sheetnames) > 0:
        wb.remove(wb[wb.sheetnames[0]])
    
    # Always create Welcome Guide first
    welcome = wb.create_sheet("Welcome Guide")
    welcome['A1'] = "🌟 Welcome to Your Life & Budget Dashboard 🌟"
    welcome['A1'].font = Font(size=24, bold=True, color='FFFFFF')
    welcome['A1'].fill = PatternFill(start_color='6f42c1', end_color='6f42c1', fill_type='solid')
    welcome['A1'].alignment = Alignment(horizontal='center', vertical='center')
    welcome.row_dimensions[1].height = 40
    
    # Add subtitle
    welcome['A2'] = "Track finances, build habits, and organize your life with ease."
    welcome['A2'].font = Font(size=14, italic=True, color='595959')
    welcome['A2'].alignment = Alignment(horizontal='center')
    welcome.row_dimensions[2].height = 25
    
    # Add some basic instructions
    welcome['A4'] = "How to Get Started:"
    welcome['A5'] = "1. Explore each tab to see what's available."
    welcome['A6'] = "2. Fill in your data in the respective trackers."
    welcome['A7'] = "3. The Dashboard tab will update automatically."
    welcome['A8'] = "4. For AI analysis, save the file and upload it in the web app."
    
    # Add footer
    welcome['A10'] = "Happy tracking! ✨"
    welcome['A10'].font = Font(size=12, italic=True, color='595959')
    welcome['A10'].alignment = Alignment(horizontal='center')
    
    # Set column widths
    welcome.column_dimensions['A'].width = 4
    welcome.column_dimensions['B'].width = 30
    welcome.column_dimensions['C'].width = 60
    welcome.column_dimensions['D'].width = 4
    
    # Save the workbook
    return wb
