# LAB14 Excel to SQLite Importer

A web-based application for importing Excel files into SQLite databases with interactive column mapping and data type handling.

## Features

- **Excel Import**: Import Excel files (.xlsx) into SQLite databases
- **Interactive Mapping**: Map Excel columns to database columns with data type selection
- **Data Types**: Support for TEXT, INTEGER, REAL, CURRENCY, and DATE types
- **Web Interface**: Modern web-based interface accessible via browser
- **Export Options**: Export data to Excel or CSV format
- **Company Support**: Company-specific data handling for LAB14 companies
- **Database Management**: View, edit, and manage database tables

## Quick Start

### For Users (Windows)
1. Go to [Releases](https://github.com/yourusername/lab14-excel-importer/releases)
2. Download the latest `LAB14_Excel_Importer_Windows.zip`
3. Extract and run `run_importer.bat`
4. Open http://localhost:5666 in your browser

### For Developers
```bash
# Clone the repository
git clone https://github.com/yourusername/lab14-excel-importer.git
cd lab14-excel-importer

# Install dependencies
pip install -r requirements.txt

# Run the application
python importxl_web.py
```

## Building Executables

### Windows Executable
The Windows executable is automatically built using GitHub Actions. To trigger a build:

1. Push changes to the `main` branch
2. GitHub Actions will automatically build the Windows executable
3. Download the artifact from the Actions tab

### Manual Build (Windows)
```cmd
# Install dependencies
pip install -r requirements.txt

# Build executable
pyinstaller importxl.spec
```

## Data Types

### Currency
- **Storage**: Values stored as pennies (cents) in database
- **Display**: Automatically converted to dollars with 2 decimal places
- **Export**: Formatted as currency in Excel exports

### Date
- **Storage**: Stored as YYYY-MM-DD format in database
- **Display**: Shown as MM/DD/YYYY in web interface
- **Import**: Supports German date format (DD.MM.YYYY)

## Supported Companies
- HIMT
- SPECSGROUP
- NANOSURF
- NANOSCRIBE
- NOTION SYSTEMS
- GENISYS
- 40-30
- AMCOSS
- MUTO TECHNOLOGIES

## Security
- Application runs on localhost only (127.0.0.1)
- No external network access
- Secure local-only operation

## Requirements
- Python 3.8+
- pandas, openpyxl, xlsxwriter, flask
- Windows 7/8/10/11 (for executable)

## License
[Your License Here] 