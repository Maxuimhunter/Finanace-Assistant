# 💰 Finance Budget Script - Your Ultimate Life Dashboard

Hey there! 👋 This is basically my all-in-one life dashboard that I built to help keep track of... well, everything! 😅

It's designed to help you:
- **Stay on top of your money game** 💳 - Track income, expenses, savings, and all that adulting stuff
- **Auto-magically process bank statements** 📄 - Extract transactions from Monzo PDFs (still working on this one!)
- **Make your data look pretty** 📊 - Generate cool charts and detailed Excel reports
- **Get your life together** 📅 - Track habits, plan meals, cleaning schedules, and personal goals
- **Get AI-powered advice** 🤖 - Smart financial recommendations to help you level up your money game

## 🚀 What It Actually Does
- **Bank Statement Parser** 🏦: Automatically extracts transactions from Monzo PDF statements *(still being worked on)*
- **Expense Tracking** 💸: Categorize and monitor where your money's going (spoiler: it's probably food 🍕)
- **Budget vs Actual** 📈: Compare what you planned to spend vs what you actually spent (oops 😅)
- **Subscription Manager** 📱: Keep track of all those monthly subscriptions that keep adding up
- **Debt Management** 💳: Track who you owe, who owes you, upcoming bills, and net debt position
- **Multi-format Export** 📤: Export to Excel, PDF, or whatever format you prefer

### Making Things Look Pretty 🎨
- **Interactive Charts** 📊: Line charts for trends, pie charts for seeing where your money goes, bar charts for comparisons
- **Financial Dashboard** 💰: Key metrics, savings rate, expense breakdown - all the important numbers
- **PDF Reports** 📄: Professional-looking reports to impress yourself (or your parents 👀)
- **Excel Wizardry** 📈: Automated Excel workbooks with multiple sheets and fancy charts

### Life Organization Stuff 📝
- **Habit Tracker** ✅: Monitor daily habits and personal development goals
- **Meal Planning** 🍳: Organize weekly meal schedules and grocery lists
- **Cleaning Schedule** 🧹: Track household maintenance tasks (because adulting is hard)
- **Goal Setting** 🎯: Set and monitor personal and financial objectives
- **Debt Tracking** 💳: Manage who you owe and who owes you, with visual pie charts

## 🛠️ Tech Stack (The Nerdy Stuff)
### Main Technologies 💻
- **Frontend**: Streamlit (basically Python magic for web apps 🪄)
- **Data Processing**: Pandas & NumPy (for making sense of all those numbers 🔢)
- **Excel Wizardry**: openpyxl (making spreadsheets look professional 📊)
- **PDF Stuff**: PyPDF2 & ReportLab (for dealing with PDFs 📄)
- **Making Things Pretty**: Matplotlib, Seaborn, Plotly (charts and graphs 📈)
- **AI Magic**: Ollama (for smart financial advice 🤖)

### Important Libraries 📚
- `streamlit` - Web app framework (makes everything look cool 🌟)
- `pandas` - Data manipulation (basically Excel on steroids 💪)
- `openpyxl` - Excel file creation (making spreadsheets fancy ✨)
- `PyPDF2` - PDF text extraction (reading bank statements 🏦)
- `reportlab` - PDF generation (creating reports 📋)
- `matplotlib` & `seaborn` - Making charts look good 📊
- `plotly` - Interactive charts (the fancy ones 🎨)
- `fpdf` - Another PDF tool (because why not? 📄)
- `ollama` - AI stuff (making the app smarter 🧠)

## 📁 How It's Organized

```
Finance Budget Script/Test Site/
├── generator.py                 # Main Streamlit application
├── enhance_budget_tracker.py    # Enhanced Excel template generator
├── debug_pdf_parser.py         # PDF parsing debugging tools
├── test_new_parser.py          # PDF parser testing
├── Enhanced_Budget_Tracker.xlsx # Sample Excel output
├── Monzo_bank_statement_*.pdf  # Sample bank statements
├── Best Version/               # Latest stable version
├── OG/                         # Original versions archive
│   ├── v2/ through v10/       # Version history (10 iterations)
│   ├── Best/                  # Best previous version
│   └── backup/                # Backup versions
├── Temp/                       # Temporary test files
└── Template/                   # Excel templates
```

## 🚀 How to Actually Use This Thing

### What You Need First 📋
1. Python 3.8 or higher (the newer the better!)
2. Virtual environment (trust me, it'll save you headaches later)
3. Required Python packages (see Installation below)

### Getting It Set Up 🔧

1. **Navigate to the project**:
   ```bash
   cd "/Users/anthonygathukia/Desktop/Me/Finance Folder's/Finance Budget Script/Test Site"
   ```

2. **Activate the virtual environment** (if you're using .venv):
   ```bash
   source .venv/bin/activate  # On macOS/Linux
   ```

3. **Install all the things**:
   ```bash
   pip install streamlit pandas numpy openpyxl PyPDF2 reportlab matplotlib seaborn plotly fpdf ollama pillow
   ```

### Running the App 🏃‍♂️

1. **Start Streamlit**:
   ```bash
   streamlit run generator.py
   ```

2. **Open your browser** and go to `http://localhost:8501`

3. **Voilà!** 🎉 Your dashboard should be running!

### How to Actually Use It 🤔

1. **Upload Bank Statements** 🏦: 
   - Upload your Monzo PDF statements and let the app do its magic
   - The parser will figure out dates, descriptions, amounts, and categories

2. **Generate Excel Reports** 📊:
   - Create awesome Excel workbooks with multiple sheets
   - Includes charts, summaries, and detailed transaction logs

3. **Check Out Your Dashboard** 💰:
   - Interactive charts showing where your money's going
   - Key metrics and financial health indicators
   - Budget vs actual comparisons (prepare for surprises 😅)

4. **Export Your Stuff** 📤:
   - Generate PDF reports with financial insights
   - Download Excel files for offline analysis

## 🔧 The Magic Behind It All

### Bank Statement Processing 🏦
1. **PDF Extraction**: Uses PyPDF2 to grab text from Monzo bank statements
2. **Transaction Parsing**: Regex patterns find transaction data (date, description, amount)
3. **Data Cleaning**: Filters out the junk and standardizes everything
4. **Categorization**: Automatically sorts transactions based on what they are

### Excel Report Generation 📈
1. **Template Creation**: Uses openpyxl to create structured Excel workbooks
2. **Data Population**: Fills multiple sheets with financial data and analysis
3. **Chart Generation**: Creates various chart types (line, pie, bar) for visualization
4. **Styling**: Makes it look professional with colors and formatting

### AI-Powered Insights 🤖
1. **Data Analysis**: Analyzes your spending patterns and financial trends
2. **Recommendation Engine**: Gives you personalized financial advice
3. **Report Generation**: Creates narrative insights based on your data

## 🐛 Debug History & My Development Journey

### Version Evolution (Like, 10 Major Updates!)
This project has been through A LOT - we're talking 10 major versions here:

- **v1-v3**: The baby days 🍼 - Basic Streamlit interface with simple expense tracking
- **v4-v6**: Getting fancy ✨ - Enhanced Excel integration and chart generation
- **v7-v8**: PDF struggles 📄 - PDF parsing capabilities and bank statement processing
- **v9-v10**: AI magic 🤖 - AI integration, advanced analytics, and professional UI

### Major Debugging Battles ⚔️

#### PDF Parser Development 🏦
- **The Problem**: Monzo PDF statements are like, super complicated
- **The Solution**: Developed multiple parsing strategies with regex patterns
- **The Evidence**: `debug_pdf_parser.py`, `test_new_parser.py` (so many test files 😅)

#### Excel Chart Integration 📊
- **The Problem**: Making charts that don't look like they're from 1995
- **The Solution**: openpyxl chart generation with custom formatting
- **The Result**: Actually professional-looking financial dashboards

#### Data Processing Pipeline 🔄
- **The Problem**: Handling all the weird transaction formats and edge cases
- **The Solution**: Robust data cleaning and validation processes
- **The Feature**: Automatic categorization and error handling (finally!)

### Testing & Validation 🧪
- **Debug Tools**: Comprehensive debugging utilities for PDF parsing
- **Test Files**: Like, a million test Excel files for validation
- **Version Control**: Kept all the old versions in `OG/` directory (hoarder much? 😅)

## 📊 Recent Changes & Cool New Stuff

### Latest Features (Best Version)
- **Enhanced UI**: Modern, responsive interface with custom CSS (so pretty! 🌟)
- **AI Integration**: Ollama-powered financial insights and recommendations
- **Advanced Charts**: 3D charts, enhanced styling, and interactive elements
- **PDF Reports**: Professional report generation with custom layouts
- **Bank Statement Automation**: Improved Monzo PDF parsing accuracy

### Performance Improvements ⚡
- **Faster Processing**: Optimized PDF parsing algorithms
- **Better Memory Management**: Efficient data handling for large datasets
- **Enhanced Error Handling**: Robust error recovery and user feedback

### User Experience Enhancements ✨
- **Intuitive Navigation**: Clear section organization and flow
- **Visual Feedback**: Progress indicators and status messages
- **Mobile Responsive**: Works on your phone too! 📱

## 🔮 What's Next? (Future Plans)

### Planned Features 🚀
- **Multi-Bank Support**: Support for other bank statement formats (not just Monzo!)
- **Advanced Analytics**: Machine learning for spending predictions (crystal ball stuff 🔮)
- **Mobile App**: Native mobile application (because why not? 📱)
- **Cloud Integration**: Sync data across devices
- **Budget Templates**: Pre-built budget templates for different lifestyles

### Technical Improvements 🛠️
- **Database Integration**: Persistent data storage
- **API Development**: RESTful API for third-party integrations
- **Security Enhancements**: User authentication and data encryption
- **Performance Optimization**: Faster processing and real-time updates

## 📝 Pro Tips

1. **Regular Updates**: Update your financial data weekly for best insights
2. **Categorization**: Review and adjust automatic categorizations (AI isn't perfect 🤷‍♂️)
3. **Goal Setting**: Set realistic financial goals and track progress
4. **Report Review**: Monthly review of generated reports and insights

## 🤝 Contributing

This is basically my personal finance project that I've been working on forever! 😅 It's focused on comprehensive life management and shows off some advanced Python development, financial analysis, and modern web application design.

Feel free to check it out and maybe get some ideas for your own projects! 🚀

## 📄 License

Personal use project - built for individual financial management and life organization. Basically, don't steal my code but feel free to learn from it! 😊

---

**Created by**: Anthony Gathukia (that's me! 👋)
**Last Updated**: January 2026
**Version**: 10+ (Best Version)
**Technology**: Python, Streamlit, Excel Integration, AI-Powered Analytics
**Age**: Born in 2006, currently in my second year of Uni 🎓
**GitHub**: https://github.com/Maxuimhunter/Finanace-Assistant
