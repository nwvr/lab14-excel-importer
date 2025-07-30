# LAB14 Excel to SQLite Importer - Windows Distribution

## Overview
This is a standalone Windows executable for the LAB14 Excel to SQLite Importer application. It allows you to import Excel files into SQLite databases with interactive column mapping and data type handling.

## Features
- Import Excel files (.xlsx) into SQLite databases
- Interactive column mapping with data type selection
- Support for currency and date data types
- Web-based interface accessible via browser
- Export data to Excel or CSV format
- Company-specific data handling
- Database viewing and editing capabilities

## Installation

### Option 1: Ready-to-Run Distribution
1. Download the `LAB14_Excel_Importer` folder
2. Extract all files to a directory of your choice
3. Double-click `run_importer.bat` to start the application
4. Open your web browser and go to: http://localhost:5666

### Option 2: Build from Source
If you need to build the executable yourself:

1. **Install Python 3.8+** from https://python.org
2. **Install required packages:**
   ```cmd
   pip install pyinstaller pandas openpyxl xlsxwriter flask
   ```
3. **Build the executable:**
   ```cmd
   pyinstaller importxl.spec
   ```
4. **Find the executable in:** `dist/LAB14_Excel_Importer/`

## Usage

### Starting the Application
1. Run `run_importer.bat` or double-click `LAB14_Excel_Importer.exe`
2. The application will start a web server on port 5666
3. Open your web browser and navigate to: http://localhost:5666

### Using the Web Interface

#### 1. File Selection
- **Excel File**: Select an existing Excel file or upload a new one
- **Database**: Choose an existing SQLite database or create a new one
- **Company**: Select the company from the dropdown list

#### 2. Table Selection
- Choose an existing table or create a new one
- The system will show available tables in the database

#### 3. Column Mapping
- Map Excel columns to database columns
- Choose data types for new columns:
  - **TEXT**: Text data
  - **INTEGER**: Whole numbers
  - **REAL**: Decimal numbers
  - **CURRENCY**: Money values (stored as pennies)
  - **DATE**: Date values (stored as YYYY-MM-DD)

#### 4. Data Import
- Review your mapping choices
- Click "Import Data" to complete the process
- View success/error messages

### Database Management

#### Viewing Data
- Click "View Database" to see all tables
- Click on a table name to view its contents
- Use the interactive table with sorting and scrolling

#### Exporting Data
- Export tables to Excel (.xlsx) format
- Export tables to CSV format
- Currency values are automatically converted to dollars for export

#### Editing Database
- Rename columns using the [Edit] links
- Delete tables if needed
- Start over with a new import

## File Structure
```
LAB14_Excel_Importer/
├── LAB14_Excel_Importer.exe    # Main executable
├── run_importer.bat            # Windows batch file to run
├── static/                     # Web assets (logo, etc.)
├── Data/                       # Excel files directory
├── config/                     # Configuration files
└── [various .dll files]       # Required libraries
```

## Troubleshooting

### Common Issues

#### 1. "Port 5666 is already in use"
- Close any other applications using port 5666
- Or modify the port in the source code and rebuild

#### 2. "Cannot find Excel files"
- Place your Excel files in the `Data/` folder
- Or use the upload feature in the web interface

#### 3. "Database errors"
- Ensure you have write permissions in the directory
- Check that the database file isn't locked by another application

#### 4. "Import fails with errors"
- Check the console output for detailed error messages
- Ensure Excel files are not corrupted
- Verify column names don't contain special characters

### Debug Mode
To run with debug information:
1. Open Command Prompt
2. Navigate to the executable directory
3. Run: `LAB14_Excel_Importer.exe --debug`

## Data Types

### Currency
- **Storage**: Values are stored as pennies (cents) in the database
- **Display**: Automatically converted to dollars with 2 decimal places
- **Export**: Formatted as currency in Excel exports

### Date
- **Storage**: Stored as YYYY-MM-DD format in database
- **Display**: Shown as MM/DD/YYYY in web interface
- **Import**: Supports German date format (DD.MM.YYYY)

### Text
- **Storage**: Stored as TEXT in database
- **Handling**: Preserves original formatting

### Numbers
- **INTEGER**: Whole numbers only
- **REAL**: Decimal numbers with precision

## Company Support
The application supports the following companies:
- HIMT
- SPECSGROUP
- NANOSURF
- NANOSCRIBE
- NOTION SYSTEMS
- GENISYS
- 40-30
- AMCOSS
- MUTO TECHNOLOGIES

Each import automatically adds a `LAB14COMPANY` column with the selected company name.

## Technical Details

### Requirements
- Windows 7/8/10/11 (64-bit)
- No additional software installation required
- Self-contained executable with all dependencies

### Port Configuration
- Default port: 5666
- Access via: http://localhost:5666
- Can be modified in source code if needed

### Database Format
- SQLite 3.x database files
- Compatible with standard SQLite tools
- Can be opened with DB Browser for SQLite

## Support
For technical support or issues:
1. Check the console output for error messages
2. Ensure all files are in the correct directories
3. Verify network/firewall settings allow local connections

## Version Information
- Application: LAB14 Excel to SQLite Importer
- Version: 1.0
- Build Date: [Current Date]
- Python Version: 3.8+
- Dependencies: pandas, openpyxl, xlsxwriter, flask, sqlite3 