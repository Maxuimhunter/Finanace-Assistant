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
- **Subscription Tracker** 🔄: Dedicated subscription management with billing cycles and auto-renewal tracking
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
- **Subscription Optimization** 🔄: Track all recurring payments and identify cost-saving opportunities

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
│   ├── v2/ through v11/       # Version history (11 iterations)
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
   - **NEW**: Select from enhanced sections including Debt Tracker and Subscription Tracker

3. **Check Out Your Dashboard** 💰:
   - Interactive charts showing where your money's going
   - Key metrics and financial health indicators
   - Budget vs actual comparisons (prepare for surprises 😅)

4. **Get AI Insights** 🤖:
   - Upload your filled Excel file for AI-powered analysis
   - Get personalized recommendations for debt management and subscription optimization
   - **NEW**: Dedicated analysis for Debt Tracker and Subscription Tracker data
   - **NEW**: Ask custom questions before generating your report
   - **NEW**: Select multiple quick questions for comprehensive analysis
   - **NEW**: Get all your questions answered in a dedicated report section

5. **Export Your Stuff** 📤:
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
4. **NEW**: Debt and Subscription analysis with optimization recommendations
5. **NEW**: Interactive Question System:
   - Custom question input for personalized inquiries
   - Multi-select quick questions for comprehensive analysis
   - Dedicated "Your Questions Answered" section in reports
   - AI responses based on your actual financial data

## 📋 Complete Update History & Changelog

### Version 11+ (January 2026) - The Subscription & Debt Era 🔄💳

#### 🆕 Major New Features
- **🔄 Subscription Tracker**: Dedicated sheet for managing recurring subscriptions
  - Track service names, amounts, billing cycles, next payment dates
  - Auto-renewal tracking and status management
  - Category-based organization (Entertainment, Music, Software, etc.)
  - AI-powered subscription optimization recommendations
- **💳 Debt Tracker**: Comprehensive debt management system
  - Track who you owe and who owes you
  - Priority management and due date tracking
  - Net debt position calculations
  - Visual debt distribution analysis
- **💬 Enhanced Question System**: Revolutionary AI interaction features
  - **Custom Questions**: Text area for multiple personalized questions (one per line)
  - **Multi-Select Quick Questions**: Checkboxes for selecting multiple predefined questions
  - **Smart Question Integration**: Questions are included in AI analysis prompt
  - **Dedicated Report Section**: "Your Questions Answered" section with personalized responses
  - **Improved UX**: Questions moved to top, right after file upload for better workflow

#### 🎨 UI/UX Enhancements
- **📋 Enhanced Section Selection**: Added emojis to all section checkboxes for better visual appeal
  - 💰 Financial section with money emojis
  - 🏋️ Health section with fitness emojis  
  - 🏠 Life Organization section with home emojis
- **🔄 Improved Organization**: Removed redundant "Monthly Purchases" sheet
- **📊 Better Sheet Ordering**: Logical flow of sheets in generated Excel files

#### 🤖 AI Integration Improvements
- **🧠 Enhanced AI Categories**: Added "Debt" and "Subscriptions" to AI analysis options
- **📈 Expanded AI Prompts**: Dedicated analysis sections for debt and subscription optimization
- **🔍 Smarter Insights**: AI can now provide specific recommendations for:
  - Debt repayment strategies (avalanche vs snowball method)
  - Subscription cost optimization and cancellation opportunities
  - Budget allocation considering debt obligations
- **💬 Revolutionary Question System**: 
  - Questions are processed and integrated into the AI analysis prompt
  - AI provides detailed, personalized answers based on actual user data
  - Multiple questions can be asked and answered in a single report
  - Questions are addressed in a dedicated section with specific recommendations

#### 🔧 Technical Improvements
- **🐛 Bug Fixes**: Resolved "monthly_purchases is not defined" error
- **📦 Import Fixes**: Added missing ollama import for AI functionality
- **🔄 Sheet Processing**: Enhanced sheet creation and mapping logic
- **⚡ Performance**: Optimized Excel template generation

### Version 10 (Late 2025) - The Debt Revolution 💳

#### 🆕 Major New Features
- **💳 Debt Tracker Feature**: Complete debt management system
  - Track who you owe and who owes you
  - Visual pie charts for debt distribution
  - Priority management and due date tracking
  - Net debt position calculations

#### 🔧 Technical Improvements
- **🛠️ Code Refactoring**: Simplified Excel generation function for better stability
- **📋 Enhanced UI**: Improved user interface with better error handling
- **🔧 Syntax Fixes**: Fixed indentation and syntax errors in generator.py
- **🐛 Major Bug Fix**: Resolved persistent Excel file corruption issues

### Version 9 (Mid 2025) - AI Magic Era 🤖

#### 🆕 Major New Features
- **🤖 AI Integration**: Ollama-powered financial insights and recommendations
- **📊 Advanced Analytics**: Machine learning-powered spending pattern analysis
- **📈 Enhanced Charts**: 3D charts, enhanced styling, and interactive elements
- **📄 PDF Reports**: Professional report generation with custom layouts

#### 🔧 Technical Improvements
- **🏦 Bank Statement Automation**: Improved Monzo PDF parsing accuracy
- **⚡ Performance**: Faster processing and better memory management
- **📱 Mobile Responsive**: Works on your phone too!

### Version 8 (Early 2025) - The PDF Struggles 📄

#### 🆕 Major New Features
- **🏦 Bank Statement Processing**: Monzo PDF statement parsing capabilities
- **📊 Enhanced Charts**: Better visual representations of financial data
- **🔍 Debug Tools**: Comprehensive debugging utilities for PDF parsing

#### 🔧 Technical Improvements
- **📋 Data Pipeline**: Robust data cleaning and validation processes
- **🧪 Testing**: Multiple test files and validation tools
- **🔄 Error Handling**: Better error recovery and user feedback

### Version 7 (Late 2025) - Chart Generation Era 📊

#### 🆕 Major New Features
- **📊 Excel Chart Integration**: Professional-looking financial dashboards
- **🎨 Visual Enhancements**: Custom formatting and styling options
- **📈 Multiple Chart Types**: Line, pie, and bar charts for different data views

#### 🔧 Technical Improvements
- **📦 openpyxl Integration**: Advanced Excel manipulation capabilities
- **🎯 Data Visualization**: Better ways to see where your money goes

### Version 6 (Mid 2025) - Excel Wizardry 📈

#### 🆕 Major New Features
- **📊 Excel Template Generation**: Automated Excel workbook creation
- **📋 Multiple Sheets**: Organized data across different tabs
- **🎨 Professional Formatting**: Colors, fonts, and styling

#### 🔧 Technical Improvements
- **📦 openpyxl Library**: Advanced Excel file manipulation
- **🔄 Template System**: Reusable Excel templates

### Version 5 (July 2025) - The Organization Era 📅

#### 🆕 Major New Features
- **📅 Life Organization**: Meal planning, cleaning schedules, habit tracking
- **✅ Habit Tracker**: Daily habit monitoring and goal setting
- **🍳 Meal Planning**: Weekly meal schedules and grocery lists
- **🧹 Cleaning Schedule**: Household maintenance task tracking

#### 🔧 Technical Improvements
- **📋 Expanded Scope**: Beyond just finances to full life management
- **🎯 Goal Setting**: Personal and financial objective tracking

### Version 4 (June 2025) - Enhanced Analytics 📈

#### 🆕 Major New Features
- **📊 Budget vs Actual**: Compare planned vs actual spending
- **💰 Savings Tracking**: Monitor savings goals and progress
- **📈 Investment Tracking**: Stock portfolio management
- **🎯 Financial Goals**: Set and track financial objectives

#### 🔧 Technical Improvements
- **📊 Data Analysis**: Better financial insights and metrics
- **💡 Recommendations**: Personalized financial advice

### Version 3 (May 2025) - The First Steps 🍼

#### 🆕 Major New Features
- **💸 Expense Tracking**: Basic expense categorization and monitoring
- **📊 Simple Charts**: Basic visual representations of spending
- **📋 Categories**: Automatic transaction categorization
- **💰 Income Tracking**: Monitor multiple income sources

#### 🔧 Technical Improvements
- **📊 pandas Integration**: Better data manipulation
- **🎨 Basic UI**: Simple Streamlit interface

### Version 2 (April 2025) - The Beginning 🌱

#### 🆕 Major New Features
- **🏦 Basic Bank Statement Parsing**: Simple PDF text extraction
- **💸 Manual Expense Entry**: Basic expense tracking functionality
- **📊 Simple Dashboard**: Basic financial overview
- **📋 Data Export**: Export data to CSV format

#### 🔧 Technical Improvements
- **📦 Basic Libraries**: Initial Streamlit and pandas setup
- **🔧 Foundation**: Core application structure

### Version 1 (January 2025) - The Concept 💡

#### 🆕 Initial Features
- **📝 Basic Idea**: Concept for personal finance management
- **🎯 Planning**: Initial design and feature planning
- **📦 Setup**: Project structure and basic setup

#### 🔧 Technical Foundation
- **🐍 Python**: Decision to use Python for development
- **🌐 Web App**: Decision to use Streamlit for interface

#### 📅 Project Timeline
- **January - April 2025**: 📝 **Idea Phase** - Concept development and planning
- **May - July 2025**: 🚀 **First Development** - Initial prototype and basic features (v2-v3)
- **July - September 2025**: 📊 **Enhancement Phase** - Analytics and organization features (v4-v5)
- **September - November 2025**: 📈 **Excel Integration** - Advanced charts and templates (v6-v7)
- **November 2025 - January 2026**: 🤖 **AI & Polish** - AI integration and final refinements (v8-v11)

## 🐛 Debug History & My Development Journey

### Major Debugging Battles ⚔️

#### Excel File Corruption Crisis 📊
- **The Problem**: Persistent Excel file corruption errors preventing file generation
- **The Investigation**: Systematic debugging by disabling features one by one
- **The Root Cause**: Complex Excel generation with advanced formulas and charts
- **The Solution**: Replaced with simplified, robust Excel template function
- **The Evidence**: Created `generator_simple.py` and `generator_fixed.py` for testing

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

#### Monthly Purchases Removal 🗑️
- **The Problem**: Redundant "Monthly Purchases" sheet causing confusion
- **The Solution**: Removed the sheet and enhanced Subscription Tracker
- **The Result**: Cleaner, more focused interface

#### Ollama Integration Issues 🤖
- **The Problem**: Missing ollama import causing AI features to fail
- **The Solution**: Added proper import and category mapping
- **The Result**: Working AI insights for all data types

### Testing & Validation 🧪
- **Debug Tools**: Comprehensive debugging utilities for PDF parsing
- **Test Files**: Like, a million test Excel files for validation
- **Version Control**: Kept all the old versions in `OG/` directory (hoarder much? 😅)

## 📊 Latest Features (Current Version)

### 🆕 New in Version 11+ (January 2026)
- **🔄 Subscription Tracker**: Complete subscription management with billing cycles
- **💳 Enhanced Debt Tracker**: Improved debt management with AI insights
- **📋 Emoji UI**: Visual section selection with intuitive emojis
- **🤖 Expanded AI Analysis**: Dedicated debt and subscription optimization
- **💬 Revolutionary Question System**: 
  - Ask multiple custom questions before report generation
  - Select multiple quick questions via checkboxes
  - Get personalized answers in a dedicated report section
  - Questions integrated into AI analysis for better responses
- **🔧 Bug Fixes**: Resolved all major stability issues

### 🎨 User Experience Enhancements
- **📱 Mobile Responsive**: Works perfectly on all devices
- **🎯 Intuitive Navigation**: Clear section organization with visual cues
- **⚡ Fast Performance**: Optimized for speed and reliability
- **🔍 Smart Defaults**: Intelligent default selections for new users

## 🔮 What's Next? (Future Plans)

*For detailed technical roadmap and current issues, see the **Known Issues & Future Improvements** section above!*

### 🚀 Immediate Priorities
1. **� Bug Fixes**: Resolve PDF download and AI data reading issues
2. **🏦 Bank Statement Enhancement**: Improve parsing accuracy and expand bank support
3. **☁️ Database Integration**: Begin SQL implementation for data persistence

### 🎯 Long-term Vision
Transform this into a comprehensive financial management platform with:
- Seamless bank integration
- AI-powered financial coaching
- Multi-device synchronization
- Advanced analytics and predictions

## 📝 Pro Tips

1. **Regular Updates**: Update your financial data weekly for best insights
2. **📊 Categorization**: Review and adjust automatic categorizations (AI isn't perfect 🤷‍♂️)
3. **🎯 Goal Setting**: Set realistic financial goals and track progress
4. **📈 Report Review**: Monthly review of generated reports and insights
5. **🔄 Subscription Audit**: Quarterly review of subscriptions for optimization opportunities
6. **💳 Debt Management**: Regular review of debt priorities and repayment strategies
7. **💬 Smart Questions**: 
   - Ask specific, detailed questions for better AI responses
   - Use multiple quick questions for comprehensive analysis
   - Review the "Your Questions Answered" section for personalized advice
   - Update questions as your financial situation changes

## 🐛 Known Issues & Current Limitations

### ⚠️ Current Issues
1. **📄 PDF Download Problems**: 
   - Minor issues when downloading reports in PDF format
   - Sometimes formatting may not render correctly
   - Workaround: Use the text download option as an alternative

2. **🤖 AI Data Reading Issues**:
   - Ollama may occasionally have difficulty reading/processing Excel file data
   - Sometimes returns generic responses instead of data-specific insights
   - Workaround: Try re-uploading the file or simplifying the data structure

### 🔧 Future Improvements & Plans

#### 🏦 Bank Statement Processing
- **Multi-Bank Support**: Expand beyond Monzo to support other bank formats
- **Enhanced Parsing**: Improve regex patterns for better transaction extraction
- **Real-time Integration**: Direct API connections to banks (future goal)
- **Statement Validation**: Better error handling for corrupted or malformed PDFs

#### ☁️ Local Cloud Storage with SQL
- **Database Integration**: Implement SQLite/PostgreSQL for persistent data storage
- **Data History**: Track financial changes over time with proper versioning
- **Multi-device Sync**: Local cloud setup for accessing data across devices
- **Backup System**: Automated backups with encryption and recovery options
- **Advanced Queries**: Enable complex financial analysis with SQL queries
- **Data Security**: Implement user authentication and data encryption

#### 🚀 Technical Roadmap
- **📱 Mobile App**: Native mobile application for on-the-go access
- **🔌 API Development**: RESTful API for third-party integrations
- **🔮 Advanced Analytics**: Machine learning for spending predictions
- **⚡ Performance**: Real-time updates and faster processing
- **🎯 Budget Templates**: Pre-built templates for different lifestyles
- **📊 Advanced Visualizations**: More interactive charts and dashboards

## 🤝 Contributing

This is basically my personal finance project that I've been working on forever! 😅 It's focused on comprehensive life management and shows off some advanced Python development, financial analysis, and modern web application design.

Feel free to check it out and maybe get some ideas for your own projects! 🚀

## 📄 License

Personal use project - built for individual financial management and life organization. Basically, don't steal my code but feel free to learn from it! 😊

---

**Created by**: Anthony Gathukia (that's me! 👋)
**Last Updated**: January 2026 (Enhanced Question System Update)
**Version**: 11+ (Enhanced with Subscription Tracker, Improved Debt Management & Revolutionary Question System)
**Technology**: Python, Streamlit, Excel Integration, AI-Powered Analytics
**Age**: Born in 2006, currently in my second year of Uni 🎓
**GitHub**: https://github.com/Maxuimhunter/Finanace-Assistant
**Recent Changes**: 
- Added Revolutionary Question System with custom and multi-select quick questions
- Enhanced AI integration with personalized question answering
- Added Subscription Tracker, enhanced Debt Management
- Improved UI/UX with better question workflow
- Fixed all major bugs and stability issues
