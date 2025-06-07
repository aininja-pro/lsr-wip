# WIP Report Automation Tool

🚀 **Automate your monthly WIP report process from 3 hours to under 5 minutes!**

This tool automatically processes your Sage accounting exports and updates your Master WIP Report while preserving all formulas, formatting, and VBA macros.

## 📋 What This Tool Does

- **Reads** GL Inquiry and WIP Worksheet exports from Sage
- **Aggregates** actual costs by job number and account type (Sub Labor 5040, Material 5030, Billing 4020)
- **Merges** actual costs with budget data from WIP Worksheet
- **Updates** your Master WIP Report Excel file without overwriting formulas
- **Creates** automatic timestamped backups before any changes
- **Generates** validation reports flagging variances over $1,000
- **Preserves** all Excel formatting, formulas, VBA macros, and print settings

## 🏗️ Setup Instructions

### Prerequisites
- Python 3.8 or higher installed
- Windows, Mac, or Linux computer
- Access to your Sage accounting exports

### Installation Steps

1. **Download or clone this project** to your computer

2. **Open terminal/command prompt** and navigate to the project folder:
   ```bash
   cd path/to/WIP-Report-Automation
   ```

3. **Create a virtual environment** (recommended):
   ```bash
   python -m venv venv
   
   # On Windows:
   venv\Scripts\activate
   
   # On Mac/Linux:
   source venv/bin/activate
   ```

4. **Install required packages**:
   ```bash
   pip install -r requirements.txt
   ```

5. **Run the application**:
   ```bash
   streamlit run src/ui/app.py
   ```

6. **Open your web browser** to the URL shown (usually http://localhost:8501)

## 📁 Required Files

You need three Excel files to run the automation:

### 1. GL Inquiry Export
- **From**: Sage accounting system
- **Contains**: Actual costs with account numbers (5040, 5030, 4020)
- **Required columns**: Job Number, Account, Description, Debit, Credit
- **File format**: .xlsx or .xls

### 2. WIP Worksheet Export  
- **From**: Sage accounting system
- **Contains**: Job budgets, status, and estimate information
- **Required columns**: Job Number, Status, Job Description, Budget amounts
- **File format**: .xlsx or .xls

### 3. Master WIP Report
- **Your existing**: Excel file with monthly tabs
- **Contains**: Formulas, formatting, and sections for 5040/5030 data
- **File format**: .xlsx or .xlsm (supports VBA macros)

## 🖥️ How to Use

### Step 1: Upload Files
1. Start the application (see setup instructions above)
2. Click "Choose file" for each of the three required files
3. Wait for green checkmarks confirming successful uploads

### Step 2: Configure Options
- **Month/Year**: Select the period to process (defaults to current month)
- **Include Closed Jobs**: Check if you want to include jobs marked as "Closed"
- **Preview Only**: Check to see results without updating files (recommended first time)
- **Create Backup**: Always keep this checked for safety

### Step 3: Process Data
1. Click "🚀 Process Data" button
2. Watch the progress bar and status messages
3. Review the data preview showing:
   - Summary statistics (total jobs, jobs with activity, large variances)
   - Merged data tables
   - GL aggregated data
   - Jobs with variances over $1,000

### Step 4: Preview Excel Changes
- Review the Excel Preview section showing:
  - Location of 5040 section (Sub Labor)
  - Location of 5030 section (Material)
  - Current data in each section

### Step 5: Apply Updates (if not preview mode)
1. Uncheck "Preview Only" if you're satisfied with the preview
2. Click "✍️ Apply Updates to Excel"
3. Wait for processing to complete
4. Download the updated files

### Step 6: Download Results
- **Updated WIP Report**: Your Excel file with new data
- **Validation Report**: Excel file listing jobs with large variances
- **Backup**: Automatically created with timestamp

## 🔧 Troubleshooting

### Common Issues

**"Column not found" errors**
- Check that your Excel files have the expected column headers
- The tool can handle minor variations (e.g., "Job No" vs "Job Number")
- Make sure files are recent exports from Sage

**"Section not found" errors**
- Verify your Master WIP Report has sections with "5040" and "5030" in the headers
- Check that you're using the correct monthly tab format

**"File locked" errors**
- Close the Master WIP Report in Excel before updating
- Make sure no other programs are using the files

**Slow performance**
- Large files (>10MB) may take longer to process
- Close other applications if your computer is running slowly

### Getting Help

1. **Check the logs**: Look in the `logs/` folder for detailed error messages
2. **Review backups**: All original files are backed up before changes
3. **Start over**: Use the "🔄 Start Over" button to reset and try again

## 📊 Data Flow Explanation

### For Non-Technical Users

Think of this tool as a smart assistant that:

1. **Reads** your accounting exports (like reading two different reports)
2. **Matches** jobs between the reports (like finding the same job in both reports)
3. **Calculates** differences between budget and actual costs (like doing math across reports)
4. **Updates** your master report (like copying numbers to the right places)
5. **Protects** your formulas (never overwrites calculations you've built)

### For Technical Users

The data processing pipeline:

```
GL Inquiry → Filter accounts (5040,5030,4020) → Aggregate by job/account type
     ↓
WIP Worksheet → Trim job numbers → Filter status → Left join with GL data
     ↓
Merged Data → Compute variances → Validate thresholds → Generate reports
     ↓
Excel Integration → Find sections → Preserve formulas → Update values → Save
```

## 🛡️ Safety Features

- **Automatic Backups**: Every file is backed up before changes
- **Formula Protection**: Never overwrites Excel formulas
- **Preview Mode**: See changes before applying them
- **Validation Reports**: Flags large variances for review
- **Error Recovery**: Detailed error messages and rollback capability
- **Audit Logging**: All operations are logged with timestamps

## 📈 Performance

- **Typical processing time**: 5-30 seconds for 50-200 jobs
- **File size limits**: Works efficiently with files up to 10MB
- **Memory usage**: Uses pandas for efficient data processing
- **Excel compatibility**: Supports .xlsx, .xlsm, and preserves VBA macros

## 🔄 Maintenance

### Backup Management
- Backups are stored in `WIP_Backups/` folder
- Files older than 30 days should be manually cleaned up
- Backup naming format: `WIP_Report_BACKUP_YYYYMMDD_HHMMSS.xlsx`

### Log Management  
- Processing logs are stored in `logs/` folder
- Logs are kept for 12 months for audit purposes
- Each log entry includes timestamp, user, files processed, and status

## 📞 Support

For technical support or feature requests:

1. Check this README first
2. Review error messages in the UI
3. Check log files in the `logs/` folder
4. Contact your IT department or the tool developer

## 🎯 Project Goals Achieved

✅ **Reduced processing time**: From 3 hours to under 5 minutes  
✅ **Preserved formulas**: Never overwrites Excel formulas or formatting  
✅ **Automated backups**: Creates safety copies before any changes  
✅ **Validation reports**: Flags variances over $1,000 for review  
✅ **User-friendly interface**: Simple web interface requiring no technical knowledge  
✅ **Error handling**: Clear messages and recovery options  
✅ **Audit trail**: Complete logging for financial compliance  

---

*This tool was designed specifically for construction company WIP reporting workflows and has been tested with real Sage accounting data exports.* 