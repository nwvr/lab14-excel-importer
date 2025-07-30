from flask import Flask, render_template_string, request, redirect, url_for, send_from_directory, send_file, session
import os
import sqlite3
import pandas as pd
import io
import datetime
import urllib.parse
import json

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'

# Global constants
DATA_DIR = 'Data'
DB_EXT = '.db'
UPLOAD_DB_DIR = 'Data'
COMPANY_CHOICES = ['HIMT', 'SPECSGROUP', 'NANOSURF', 'NANOSCRIBE', 'NOTION SYSTEMS', 'GENISYS', '40-30', 'AMCOSS', 'MUTO TECHNOLOGIES']

# Create config directory for storing user preferences
CONFIG_DIR = 'config'
if not os.path.exists(CONFIG_DIR):
    os.makedirs(CONFIG_DIR)

def get_user_type_choices_file():
    """Get the path to the user type choices file"""
    return os.path.join(CONFIG_DIR, 'user_type_choices.json')

def load_user_type_choices():
    """Load user type choices from JSON file"""
    file_path = get_user_type_choices_file()
    if os.path.exists(file_path):
        try:
            with open(file_path, 'r') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {}
    return {}

def save_user_type_choices(choices):
    """Save user type choices to JSON file"""
    file_path = get_user_type_choices_file()
    try:
        with open(file_path, 'w') as f:
            json.dump(choices, f, indent=2)
    except IOError:
        pass  # Silently fail if we can't write the file

def store_user_type_choice(column_name, chosen_type):
    """Store user's preferred type for a column"""
    choices = load_user_type_choices()
    now = datetime.datetime.now().isoformat()
    
    if column_name in choices:
        choices[column_name]['chosen_type'] = chosen_type
        choices[column_name]['count'] = choices[column_name].get('count', 0) + 1
        choices[column_name]['last_used'] = now
    else:
        choices[column_name] = {
            'chosen_type': chosen_type,
            'count': 1,
            'last_used': now
        }
    
    save_user_type_choices(choices)

def get_user_type_choice(column_name):
    """Get user's preferred type for a column"""
    choices = load_user_type_choices()
    if column_name in choices:
        return choices[column_name]['chosen_type']
    return None

STYLE = '''
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&display=swap');
body {
    background: #ebebeb;
    font-family: 'Sora', 'Inter', 'Segoe UI', Arial, sans-serif;
    color: #003d21;
    margin: 0;
    padding: 0;
    line-height: 1.6;
}
.header {
    background-color: #003d21;
    color: white;
    padding: 15px 20px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
.header-logo {
    height: 50px;
    margin-right: 15px;
}
.header-title {
    font-size: 18px;
    font-weight: normal;
}
.menu {
    display: flex;
    gap: 20px;
}
.menu-link {
    color: white;
    text-decoration: none;
    padding: 8px 16px;
    border-radius: 4px;
    transition: background-color 0.3s;
}
.menu-link:hover {
    background-color: rgba(255,255,255,0.1);
}
.container {
    max-width: 1200px;
    margin: 20px auto;
    padding: 0 20px;
}
h1, h2 {
    color: #003d21;
    margin-bottom: 20px;
}
form {
    background: white;
    padding: 25px;
    border-radius: 8px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    margin-bottom: 20px;
}
input[type="text"], input[type="file"], select {
    width: 100%;
    padding: 10px;
    margin: 8px 0;
    border: 1px solid #ddd;
    border-radius: 4px;
    box-sizing: border-box;
}
input[type="submit"], button {
    background-color: #003d21;
    color: white;
    padding: 12px 24px;
    border: none;
    border-radius: 4px;
    cursor: pointer;
    font-size: 16px;
    margin-top: 10px;
}
input[type="submit"]:hover, button:hover {
    background-color: #002a1a;
}
.message {
    padding: 15px;
    margin: 15px 0;
    border-radius: 4px;
    background-color: #d4edda;
    color: #155724;
    border: 1px solid #c3e6cb;
}
.error {
    background-color: #f8d7da;
    color: #721c24;
    border: 1px solid #f5c6cb;
}
.col-map-block {
    background: #f8f9fa;
    padding: 15px;
    margin: 10px 0;
    border-radius: 6px;
    border-left: 4px solid #003d21;
}
table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    border-radius: 8px;
    overflow: hidden;
    margin: 20px 0;
}
th, td {
    padding: 12px;
    text-align: left;
    border-bottom: 1px solid #ddd;
}
th {
    background-color: #003d21;
    color: white;
    font-weight: normal;
}
.export-buttons {
    margin: 20px 0;
}
.export-buttons a {
    display: inline-block;
    background-color: #003d21;
    color: white;
    padding: 10px 20px;
    text-decoration: none;
    border-radius: 4px;
    margin-right: 10px;
}
.export-buttons a:hover {
    background-color: #002a1a;
}
.sortable {
    cursor: pointer;
    user-select: none;
}
.sortable:hover {
    background-color: #002a1a;
}
</style>
'''

def infer_sqltype(dtype):
    if pd.api.types.is_integer_dtype(dtype):
        return 'INTEGER'
    elif pd.api.types.is_float_dtype(dtype):
        return 'REAL'
    elif pd.api.types.is_bool_dtype(dtype):
        return 'BOOLEAN'
    else:
        return 'TEXT'

def create_metadata_table(conn, table_name):
    """Create a metadata table to track column types"""
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS column_metadata (
            table_name TEXT,
            column_name TEXT,
            data_type TEXT,
            is_currency INTEGER DEFAULT 0,
            is_date INTEGER DEFAULT 0,
            PRIMARY KEY (table_name, column_name)
        )
    ''')
    conn.commit()

def store_column_metadata(conn, table_name, column_types):
    """Store column type information in metadata table"""
    cur = conn.cursor()
    create_metadata_table(conn, table_name)
    
    for col, col_type in column_types.items():
        is_currency = 1 if col_type == 'CURRENCY' else 0
        is_date = 1 if col_type == 'DATE' else 0
        cur.execute('''
            INSERT OR REPLACE INTO column_metadata (table_name, column_name, data_type, is_currency, is_date)
            VALUES (?, ?, ?, ?, ?)
        ''', (table_name, col, col_type, is_currency, is_date))
    conn.commit()

def get_column_metadata(conn, table_name):
    """Get column metadata for a table"""
    cur = conn.cursor()
    cur.execute('''
        SELECT column_name, data_type, is_currency, is_date 
        FROM column_metadata 
        WHERE table_name = ?
    ''', (table_name,))
    return {row[0]: {'type': row[1], 'is_currency': bool(row[2]), 'is_date': bool(row[3])} for row in cur.fetchall()}

def _predict_column_type_base(df, column_name):
    """Predict if a column is currency or date based on content analysis"""
    col_data = df[column_name].dropna()
    if len(col_data) == 0:
        return 'TEXT'  # Default for empty columns
    
    # Check for date patterns first (before currency)
    date_keywords = ['date', 'time', 'created', 'modified', 'updated', 'start', 'end', 
                    'due', 'deadline', 'birth', 'anniversary', 'expiry', 'valid', 'period',
                    'datum', 'erstellt', 'geändert', 'schluss', 'phase']
    col_lower = column_name.lower()
    if any(keyword in col_lower for keyword in date_keywords):
        # Test if the data actually looks like dates
        try:
            sample_data = col_data.head(10)
            # Try German date format first
            parsed_dates = pd.to_datetime(sample_data, format='%d.%m.%Y', errors='coerce')
            if parsed_dates.notna().sum() / len(sample_data) > 0.5:
                return 'DATE'
            # Try other date formats
            parsed_dates = pd.to_datetime(sample_data, errors='coerce')
            if parsed_dates.notna().sum() / len(sample_data) > 0.5:
                return 'DATE'
        except:
            pass
    
    # Check for currency patterns
    if df[column_name].dtype in ['int64', 'float64']:
        currency_keywords = ['price', 'cost', 'amount', 'value', 'revenue', 'sales', 'budget', 
                           'fee', 'charge', 'payment', 'total', 'sum', 'money', 'dollar', 'euro', 
                           'currency', 'cash', 'income', 'expense', 'profit', 'loss', 'revenue',
                           'betrag', 'preis', 'kosten', 'umsatz']
        col_lower = column_name.lower()
        if any(keyword in col_lower for keyword in currency_keywords):
            return 'CURRENCY'
        if len(col_data) > 0:
            min_val = col_data.min()
            max_val = col_data.max()
            if 0.01 <= min_val <= 999999 and max_val <= 999999:
                if df[column_name].dtype == 'float64' or any(val % 1 != 0 for val in col_data if isinstance(val, (int, float))):
                    return 'CURRENCY'
    
    # Check if it's a simple integer column (not date, not currency)
    if df[column_name].dtype in ['int64']:
        # If it's a small integer range, it might be an ID or count
        if len(col_data) > 0:
            min_val = col_data.min()
            max_val = col_data.max()
            if min_val >= 0 and max_val <= 999999:
                return 'INTEGER'
    
    return infer_sqltype(df[column_name].dtype)

def predict_column_type(df, column_name, conn=None):
    """Predict column type, checking user choices first"""
    user_choice = get_user_type_choice(column_name)
    if user_choice:
        return user_choice
    return _predict_column_type_base(df, column_name)

@app.route('/', methods=['GET', 'POST'])
def index():
    excel_files = [f for f in os.listdir(DATA_DIR) if f.endswith('.xlsx') and not f.startswith('~$')]
    db_files = [f for f in os.listdir('.') if f.endswith(DB_EXT)]
    message = ''
    selected_company = request.form.get('company') if request.method == 'POST' else None
    if request.method == 'POST':
        excel_choice = request.form.get('excel_file')
        db_choice = request.form.get('db_file')
        new_db_name = request.form.get('new_db_name')
        # Handle Excel upload
        if 'excel_upload' in request.files and request.files['excel_upload'].filename:
            excel_file = request.files['excel_upload']
            excel_path = os.path.join(DATA_DIR, excel_file.filename)
            excel_file.save(excel_path)
            excel_choice = excel_file.filename
        # Handle DB upload
        if 'db_upload' in request.files and request.files['db_upload'].filename:
            db_file = request.files['db_upload']
            db_path = os.path.join(UPLOAD_DB_DIR, db_file.filename)
            db_file.save(db_path)
            db_choice = db_file.filename
        if not excel_choice:
            message = 'Please select or upload an Excel file.'
        elif not db_choice and not new_db_name:
            message = 'Please select, upload, or create a database.'
        elif not selected_company:
            message = 'Please select a company.'
        else:
            # Save choices to session and go to table selection
            session['excel_choice'] = excel_choice
            session['db_choice'] = db_choice if db_choice else (new_db_name if new_db_name.endswith(DB_EXT) else new_db_name + DB_EXT)
            session['selected_company'] = selected_company
            db_path = db_choice if db_choice else (new_db_name if new_db_name.endswith(DB_EXT) else new_db_name + DB_EXT)
            return redirect(url_for('table_select', excel=excel_choice, db=db_path, company=selected_company))
    return render_template_string(STYLE + '''
    <div class="header"><img src="/static/Unknown.png" alt="LAB14 Logo" class="header-logo"><span class="header-title">Excel to SQLite Importer</span><nav class="menu"><a href="/" class="menu-link">Start Over</a><a href="/view_db" class="menu-link">View Database</a></nav></div>
    <div class="container">
    <form method="post" enctype="multipart/form-data">
        <h2>Select or upload Excel file</h2>
        <select name="excel_file">
            <option value="">-- Select Excel file --</option>
            {% for f in excel_files %}
            <option value="{{f}}">{{f}}</option>
            {% endfor %}
        </select>
        or upload: <input type="file" name="excel_upload"><br>
        <h2>Select, upload, or create SQLite DB</h2>
        <select name="db_file">
            <option value="">-- Select DB file --</option>
            {% for f in db_files %}
            <option value="{{f}}">{{f}}</option>
            {% endfor %}
        </select>
        or upload: <input type="file" name="db_upload"><br>
        or new DB name: <input type="text" name="new_db_name"><br>
        <h2>Select Company</h2>
        <select name="company">
            <option value="">-- Select Company --</option>
            {% for c in company_choices %}
            <option value="{{c}}" {% if c == selected_company %}selected{% endif %}>{{c}}</option>
            {% endfor %}
        </select>
        <input type="submit" value="Next">
    </form>
    <div class="message">{{message}}</div>
    </div>
    ''', excel_files=excel_files, db_files=db_files, message=message, company_choices=COMPANY_CHOICES, selected_company=selected_company)

@app.route('/table_select', methods=['GET', 'POST'])
def table_select():
    excel = request.args.get('excel')
    db = request.args.get('db')
    company = request.args.get('company')
    table_list = []
    message = ''
    
    # Ensure database path is correct
    if db:
        if not os.path.isabs(db):
            db_path = os.path.join('.', db)
        else:
            db_path = db
        
        if os.path.exists(db_path):
            try:
                conn = sqlite3.connect(db_path)
                cur = conn.cursor()
                cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
                table_list = [row[0] for row in cur.fetchall()]
                conn.close()
            except Exception as e:
                message = f"Could not read tables: {e}"
    
    if request.method == 'POST':
        table_choice = request.form.get('table_name')
        new_table = request.form.get('new_table_name')
        if not table_choice and not new_table:
            message = 'Please select or enter a table name.'
        else:
            if new_table:
                table_choice = new_table
            return redirect(url_for('column_mapping', excel=excel, db=db, table=table_choice, company=company))
    
    return render_template_string(STYLE + '''
    <div class="header"><img src="/static/Unknown.png" alt="LAB14 Logo" class="header-logo"><span class="header-title">Excel to SQLite Importer</span><nav class="menu"><a href="/" class="menu-link">Start Over</a><a href="/view_db" class="menu-link">View Database</a></nav></div>
    <div class="container">
    <form method="post">
        <h2>Select an existing table</h2>
        <select name="table_name">
            <option value="">-- Select Table --</option>
            {% for t in table_list %}
            <option value="{{t}}">{{t}}</option>
            {% endfor %}
        </select>
        or create new table: <input type="text" name="new_table_name"><br>
        <input type="submit" value="Next">
    </form>
    <div class="message">{{message}}</div>
    </div>
    ''', table_list=table_list, message=message)

@app.route('/column_mapping', methods=['GET', 'POST'])
def column_mapping():
    excel = request.args.get('excel')
    db = request.args.get('db')
    table = request.args.get('table')
    company = request.args.get('company')
    
    print(f"DEBUG: column_mapping called with:")
    print(f"  excel: {excel}")
    print(f"  db: {db}")
    print(f"  table: {table}")
    print(f"  company: {company}")
    
    excel_path = os.path.join(DATA_DIR, excel)
    df = pd.read_excel(excel_path, sheet_name=0)
    if company:
        df['LAB14COMPANY'] = company
    
    # Ensure database path is correct
    if not os.path.isabs(db):
        db_path = os.path.join('.', db)
    else:
        db_path = db
    
    print(f"DEBUG: Using database path: {db_path}")
    print(f"DEBUG: Table name for import: '{table}'")
    
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    # Get table columns
    cur.execute(f"PRAGMA table_info('{table}')")
    table_info = cur.fetchall()
    table_columns = [col[1] for col in table_info]
    message = ''  # Initialize message variable
    # Get column types from session (selected earlier) - moved up to avoid UnboundLocalError
    col_types = session.get('col_types', {})
    type_choices = ['TEXT', 'INTEGER', 'REAL', 'CURRENCY', 'DATE']
    
    # If table does not exist, let user select which columns to use and their types
    if not table_columns:
        inferred_types = {col: predict_column_type(df, col, conn) for col in df.columns}
        if request.method == 'POST':
            selected_cols = request.form.getlist('use_col')
            col_defs = []
            # Store column types for later use
            column_types_to_store = {}
            for col in selected_cols:
                col_type = request.form.get(f'type_{col}', inferred_types[col])
                column_types_to_store[col] = col_type
                sqltype = col_type
                if sqltype == 'CURRENCY':
                    sqltype = 'INTEGER'
                elif sqltype == 'DATE':
                    sqltype = 'TEXT'  # Store dates as text in SQLite
                col_defs.append(f'"{col}" {sqltype}')
            if not col_defs:
                message = 'Please select at least one column.'
            else:
                cur.execute(f'CREATE TABLE "{table}" ({", ".join(col_defs)})')
                conn.commit()
                # Store column metadata
                store_column_metadata(conn, table, column_types_to_store)
                table_columns = selected_cols
                table_info = [(None, col, None, None, None, None) for col in selected_cols]
                # After creation, redirect to self to trigger normal mapping flow
                conn.close()
                return redirect(url_for('column_mapping', excel=excel, db=db, table=table, company=company))
        return render_template_string(STYLE + '''
        <div class="header"><img src="/static/Unknown.png" alt="LAB14 Logo" class="header-logo"><span class="header-title">Excel to SQLite Importer</span><nav class="menu"><a href="/" class="menu-link">Start Over</a><a href="/view_db" class="menu-link">View Database</a></nav></div>
        <div class="container">
        <h1>Create Table: {{table}}</h1>
        <form method="post">
            <p>Select which columns from the Excel file to include in the new table and their data types:</p>
            {% for col in df.columns %}
            <div class="col-map-block">
                <input type="checkbox" name="use_col" value="{{col}}" id="col_{{loop.index}}" checked>
                <label for="col_{{loop.index}}"><b>{{col}}</b></label>
                <span style="margin-left:10px;">Type:
                    <select name="type_{{col}}">
                        {% for t in type_choices %}
                        <option value="{{t}}" {% if t == inferred_types[col] %}selected{% endif %}>{{t}}</option>
                        {% endfor %}
                    </select>
                </span>
                {% if inferred_types[col] in ['CURRENCY', 'DATE'] %}
                <div style="margin-top:5px;font-size:0.9em;color:#007acc;">
                    <strong>💡 Smart Prediction:</strong> Detected as {{inferred_types[col]}} based on column name and data analysis
                </div>
                {% endif %}
            </div>
            {% endfor %}
            <input type="submit" value="Create Table and Continue">
        </form>
        <div class="message">{{message}}</div>
        </div>
        ''', table=table, df=df, col_types=col_types, message=message, type_choices=type_choices, inferred_types=inferred_types)
    
    # Get all existing table columns for reference
    all_table_columns = [col[1] for col in table_info]
    
    # Get existing column types from table schema and metadata
    metadata = get_column_metadata(conn, table)
    existing_column_types = {}
    for col_info in table_info:
        col_name = col_info[1]
        col_type = col_info[2]
        # Use metadata to show CURRENCY/DATE if set
        if col_name in metadata:
            if metadata[col_name].get('is_currency'):
                col_type = 'CURRENCY'
            elif metadata[col_name].get('is_date'):
                col_type = 'DATE'
        existing_column_types[col_name] = col_type
    
    # Get inferred types for all columns (needed for template)
    inferred_types = {col: predict_column_type(df, col, conn) for col in df.columns}
    
    print(f"DEBUG: Template variables:")
    print(f"  df.columns: {list(df.columns)}")
    print(f"  all_table_columns: {all_table_columns}")
    print(f"  existing_column_types: {existing_column_types}")
    
    if request.method == 'POST':
        print(f"DEBUG: POST request received")
        mapping = {}
        message = ''  # Initialize message for POST requests
        
        # Debug: Print what we received
        print(f"DEBUG: Processing POST request for table {table}")
        print(f"DEBUG: Excel columns: {list(df.columns)}")
        print(f"DEBUG: Table columns: {all_table_columns}")
        print(f"DEBUG: Form data received:")
        for key, value in request.form.items():
            print(f"  {key}: {value}")
        
        # Check which columns exist in the table
        print(f"DEBUG: Checking which Excel columns exist in table:")
        for col in df.columns:
            exists = col in all_table_columns
            print(f"  {col}: {'EXISTS' if exists else 'NEW'}")
        
        # Process all Excel columns (both missing and existing)
        for col in df.columns:
            action = request.form.get(f'action_{col}')
            print(f"DEBUG: Column {col} -> action: {action}")
            
            # Check if this column should be existing or new
            should_be_existing = col in all_table_columns
            print(f"DEBUG: Column {col} should be existing: {should_be_existing}, but action is: {action}")
            
            # Auto-correct: If column exists in table but action is 'create', change to map to self
            if should_be_existing and action == 'create':
                print(f"DEBUG: Auto-correcting: {col} exists but action is 'create', changing to map to self")
                action = col
            
            if action == 'skip':
                print(f"DEBUG: Skipping column {col}")
                continue
            elif action == 'create':
                # Create new column
                new_col_name = request.form.get(f'rename_{col}', col) # Get new name or default
                col_type = request.form.get(f'type_{col}', 'TEXT')
                print(f"DEBUG: Creating column {new_col_name} with type {col_type}")
                
                # Handle empty column names
                if not new_col_name or new_col_name.strip() == '':
                    new_col_name = col  # Use original column name as default
                    print(f"DEBUG: Empty column name detected, using original name: {new_col_name}")
                
                sqltype = col_type
                # Handle CURRENCY and DATE types
                if sqltype == 'CURRENCY':
                    sqltype = 'INTEGER'
                elif sqltype == 'DATE':
                    sqltype = 'TEXT'  # Store dates as text in SQLite
                try:
                    # Check if column already exists before trying to create it
                    if new_col_name not in all_table_columns:
                        cur.execute(f'ALTER TABLE "{table}" ADD COLUMN "{new_col_name}" {sqltype}')
                        conn.commit()
                        mapping[col] = new_col_name # Map original to new name
                        # Store metadata for this column
                        store_column_metadata(conn, table, {new_col_name: col_type})
                        print(f"DEBUG: Created new column {new_col_name} for {col}")
                    else:
                        # Column already exists, just map to it
                        mapping[col] = new_col_name
                        print(f"DEBUG: Column {new_col_name} already exists, mapping {col} to it")
                except Exception as e:
                    error_msg = f'Failed to add column {new_col_name}: {e}'
                    message += error_msg + '<br>'
                    print(f"DEBUG: Failed to create column {new_col_name}: {e}")
                    # Still add to mapping to allow import to proceed
                    mapping[col] = new_col_name
                    print(f"DEBUG: Added {col} to mapping despite error to allow import to proceed")
            elif action == 'map_other':
                # Map to a different existing column
                target_col = request.form.get(f'map_to_{col}')
                if target_col and target_col in all_table_columns:
                    mapping[col] = target_col
                    print(f"DEBUG: Mapped {col} to existing column {target_col}")
                else:
                    message += f'Invalid mapping for column {col}<br>'
                    print(f"DEBUG: Invalid mapping for {col} to {target_col}")
            else:
                # Map to existing column (action contains the target column name)
                if action in all_table_columns:
                    mapping[col] = action
                    print(f"DEBUG: Mapped {col} to existing column {action}")
                else:
                    message += f'Invalid mapping for column {col}<br>'
                    print(f"DEBUG: Invalid mapping for {col} to {action}")
        
        print(f"DEBUG: Final mapping: {mapping}")
        print(f"DEBUG: Message: {message}")
        
        # Only proceed if we have mappings (even if there are some errors)
        if mapping:
            print(f"DEBUG: Proceeding with import - mapping has {len(mapping)} columns")
            # Prepare DataFrame for import
            import_df = df.copy()
            import_df = import_df[[c for c in mapping.keys()]]
            import_df = import_df.rename(columns=mapping)
            
            # Clean up column names - remove empty or problematic names
            problematic_cols = []
            for col in import_df.columns:
                if not col or col.strip() == '' or col.startswith('Unnamed:'):
                    problematic_cols.append(col)
            
            if problematic_cols:
                print(f"DEBUG: Found problematic columns: {problematic_cols}")
                # Remove problematic columns
                import_df = import_df.drop(columns=problematic_cols)
                print(f"DEBUG: Removed problematic columns, remaining: {list(import_df.columns)}")
            
            # Validate that we have data to import
            if import_df.empty:
                message = 'No data to import after column mapping.'
                print(f"DEBUG: DataFrame is empty after mapping")
            elif not table or table.strip() == '':
                message = 'Invalid table name specified.'
                print(f"DEBUG: Table name is empty or invalid: '{table}'")
            else:
                print(f"DEBUG: About to import {len(import_df)} rows to table '{table}'")
                print(f"DEBUG: Import columns: {list(import_df.columns)}")
                
                # Handle currency conversion for import
                column_types_to_store = {}  # Track all column types for metadata storage
                
                for col in import_df.columns:
                    # Determine if this is an existing column or a new one
                    is_existing_column = col in all_table_columns
                    
                    if is_existing_column:
                        # For existing columns, use the existing type from the database
                        col_type = existing_column_types[col]
                        # Store metadata for existing columns to ensure consistency
                        column_types_to_store[col] = col_type
                    else:
                        # For new columns, get the type from the form
                        col_type = request.form.get(f'type_{col}', 'TEXT')
                        # Get the new column name (if renamed)
                        new_col_name = request.form.get(f'rename_{col}', col)
                        # Store metadata using the new column name
                        column_types_to_store[new_col_name] = col_type
                    
                    print(f"DEBUG: Processing column {col} with type {col_type}")
                    
                    # Handle currency conversion based on column type (not just metadata)
                    if col_type == 'CURRENCY':
                        # Convert to pennies/cents (multiply by 100 and round to integer)
                        print(f"DEBUG: Converting {col} to currency (pennies)")
                        original_values = import_df[col].copy()
                        # Handle NaN values - only convert non-NaN values
                        import_df[col] = pd.to_numeric(import_df[col], errors='coerce')
                        # Only multiply non-NaN values by 100
                        mask = import_df[col].notna()
                        import_df.loc[mask, col] = import_df.loc[mask, col].round(2) * 100
                        # Convert to integer more safely
                        try:
                            import_df[col] = import_df[col].astype('Int64')  # Use nullable integer type
                        except TypeError:
                            # If Int64 fails, try regular int64 and handle NaN values
                            import_df[col] = import_df[col].fillna(0).astype('int64')
                        print(f"DEBUG: Currency conversion - Original: {original_values.head()}, Converted: {import_df[col].head()}")
                    elif col_type == 'INTEGER':
                        # Convert to integer, handling NaN values
                        print(f"DEBUG: Converting {col} to integer")
                        original_values = import_df[col].copy()
                        import_df[col] = pd.to_numeric(import_df[col], errors='coerce')
                        # Convert to integer more safely
                        try:
                            import_df[col] = import_df[col].astype('Int64')  # Use nullable integer type
                        except TypeError:
                            # If Int64 fails, try regular int64 and handle NaN values
                            import_df[col] = import_df[col].fillna(0).astype('int64')
                        print(f"DEBUG: Integer conversion - Original: {original_values.head()}, Converted: {import_df[col].head()}")
                    elif col_type == 'DATE':
                        # Convert to ISO format string
                        print(f"DEBUG: Converting {col} to date format")
                        original_values = import_df[col].copy()
                        # Try German date format first (dd.mm.yyyy), then other formats
                        import_df[col] = pd.to_datetime(import_df[col], format='%d.%m.%Y', errors='coerce')
                        # If that fails, try other common formats
                        if import_df[col].isna().all():
                            import_df[col] = pd.to_datetime(original_values, errors='coerce')
                        import_df[col] = import_df[col].dt.strftime('%Y-%m-%d')
                        print(f"DEBUG: Date conversion - Original: {original_values.head()}, Converted: {import_df[col].head()}")
                    elif col_type == 'REAL':
                        # Convert to float, handling NaN values
                        print(f"DEBUG: Converting {col} to real/float")
                        original_values = import_df[col].copy()
                        import_df[col] = pd.to_numeric(import_df[col], errors='coerce')
                        print(f"DEBUG: Real conversion - Original: {original_values.head()}, Converted: {import_df[col].head()}")
                
                print(f"DEBUG: DataFrame after type conversions:")
                print(f"  Shape: {import_df.shape}")
                print(f"  Sample data:")
                print(import_df.head())
                
                # Import the data
                try:
                    print(f"DEBUG: DataFrame before import:")
                    print(f"  Shape: {import_df.shape}")
                    print(f"  Columns: {list(import_df.columns)}")
                    print(f"  First few rows:")
                    print(import_df.head())
                    
                    # Check table structure before import
                    cur = conn.cursor()
                    cur.execute(f"PRAGMA table_info('{table}')")
                    table_info = cur.fetchall()
                    print(f"DEBUG: Table structure before import:")
                    for col_info in table_info:
                        print(f"  {col_info[1]} ({col_info[2]})")
                    
                    # Get row count before import
                    cur.execute(f"SELECT COUNT(*) FROM {table}")
                    count_before = cur.fetchone()[0]
                    print(f"DEBUG: Row count before import: {count_before}")
                    
                    # Try the import
                    import_df.to_sql(table, conn, if_exists='append', index=False)
                    conn.commit()
                    
                    # Get row count after import
                    cur.execute(f"SELECT COUNT(*) FROM {table}")
                    count_after = cur.fetchone()[0]
                    print(f"DEBUG: Row count after import: {count_after}")
                    print(f"DEBUG: Rows added: {count_after - count_before}")
                    
                    # Check if any data was actually inserted
                    if count_after == count_before:
                        print(f"DEBUG: WARNING - No rows were actually inserted!")
                        # Try to get a sample of the data to see what's wrong
                        cur.execute(f"SELECT * FROM {table} LIMIT 5")
                        sample_data = cur.fetchall()
                        print(f"DEBUG: Sample data from table: {sample_data}")
                    
                    # Store column metadata for ALL columns (both new and existing)
                    store_column_metadata(conn, table, column_types_to_store)
                    # Store user data type choices for future prediction
                    for col, col_type in column_types_to_store.items():
                        store_user_type_choice(col, col_type)
                    
                    # Show success message with any warnings
                    if message:
                        message = f'Successfully imported {len(import_df)} rows to table "{table}". Some columns could not be created: {message}'
                    else:
                        message = f'Successfully imported {len(import_df)} rows to table "{table}"'
                    
                    conn.close()
                    return redirect(url_for('view_db'))
                except Exception as e:
                    message = f'Import failed: {e}'
                    print(f"DEBUG: Import exception: {e}")
                    conn.close()
        elif not mapping:
            message = 'No columns were selected for import.'
            print(f"DEBUG: No columns were mapped")
        else:
            print(f"DEBUG: Import blocked due to errors: {message}")
        # If there are errors, continue to show the form with error messages
    
    # Prepare template variables
    type_choices = ['TEXT', 'INTEGER', 'REAL', 'CURRENCY', 'DATE']
    inferred_types = {col: predict_column_type(df, col, conn) for col in df.columns}
    
    return render_template_string(STYLE + '''
    <div class="header"><img src="/static/Unknown.png" alt="LAB14 Logo" class="header-logo"><span class="header-title">Excel to SQLite Importer</span><nav class="menu"><a href="/" class="menu-link">Start Over</a><a href="/view_db" class="menu-link">View Database</a></nav></div>
    <div class="container">
    <h1>Column Mapping for Table: {{table}}</h1>
    <form method="post">
        <h2>Map Excel Columns to Database</h2>
        <p>For each Excel column, choose how to handle it:</p>
        
        {% for col in df.columns %}
        <div class="col-map-block">
            <b>{{col}}</b><br>
            
            {% if col in all_table_columns %}
            <!-- For existing columns, default to mapping to the same column -->
            <input type="radio" name="action_{{col}}" value="{{col}}" id="map_{{col}}" checked>
            <label for="map_{{col}}">Map to existing column: <strong>{{col}}</strong> ({{existing_column_types[col]}})</label><br>
            
            <!-- Map to other existing column option -->
            <input type="radio" name="action_{{col}}" value="map_other" id="map_other_{{col}}">
            <label for="map_other_{{col}}">Map to other existing column:</label>
            <select name="map_to_{{col}}" disabled>
                <option value="">-- Select column --</option>
                {% for table_col in all_table_columns %}
                {% if table_col != col %}
                <option value="{{table_col}}">{{table_col}} ({{existing_column_types[table_col]}})</option>
                {% endif %}
                {% endfor %}
            </select><br>
            
            <!-- Create new column option (disabled by default for existing columns) -->
            <input type="radio" name="action_{{col}}" value="create" id="create_{{col}}">
            <label for="create_{{col}}">Create new column</label>
            <input type="text" name="rename_{{col}}" placeholder="New column name (optional)" style="width: 200px; margin-left: 10px;" disabled>
            <select name="type_{{col}}" disabled>
                {% for t in type_choices %}
                <option value="{{t}}" {% if t == existing_column_types[col] %}selected{% endif %}>{{t}}</option>
                {% endfor %}
            </select>
            <span style="color: #666; font-size: 0.9em;">(Type selection disabled for existing columns)</span><br>
            
            {% else %}
            <!-- For new columns, default to creating new column -->
            <input type="radio" name="action_{{col}}" value="create" id="create_{{col}}" checked>
            <label for="create_{{col}}">Create new column</label>
            <input type="text" name="rename_{{col}}" placeholder="New column name (optional)" style="width: 200px; margin-left: 10px;">
            <select name="type_{{col}}">
                {% for t in type_choices %}
                <option value="{{t}}" {% if t == inferred_types[col] %}selected{% endif %}>{{t}}</option>
                {% endfor %}
            </select><br>
            
            <!-- Map to existing column option (only if there are existing columns) -->
            {% if all_table_columns %}
            <input type="radio" name="action_{{col}}" value="map_other" id="map_other_{{col}}">
            <label for="map_other_{{col}}">Map to existing column:</label>
            <select name="map_to_{{col}}" disabled>
                <option value="">-- Select column --</option>
                {% for table_col in all_table_columns %}
                <option value="{{table_col}}">{{table_col}} ({{existing_column_types[table_col]}})</option>
                {% endfor %}
            </select><br>
            {% endif %}
            
            {% endif %}
            
            <!-- Skip option for all columns -->
            <input type="radio" name="action_{{col}}" value="skip" id="skip_{{col}}">
            <label for="skip_{{col}}">Skip this column</label><br>
            
            {% if inferred_types[col] in ['CURRENCY', 'DATE'] %}
            <div style="margin-top:5px;font-size:0.9em;color:#007acc;">
                <strong>💡 Smart Prediction:</strong> Detected as {{inferred_types[col]}} based on column name and data analysis
            </div>
            {% endif %}
        </div>
        {% endfor %}
        
        <input type="submit" value="Import Data">
    </form>
    <div class="message">{{message}}</div>
    </div>
    
    <script>
    // Enable/disable form elements based on radio button selection
    document.addEventListener('DOMContentLoaded', function() {
        const excelColumns = {{df.columns.tolist() | tojson}};
        const tableColumns = {{all_table_columns | tojson}};
        
        excelColumns.forEach(function(col) {
            const createRadio = document.getElementById('create_' + col);
            const typeSelect = document.querySelector('select[name="type_' + col + '"]');
            const mapOtherRadio = document.getElementById('map_other_' + col);
            const mapOtherSelect = document.querySelector('select[name="map_to_' + col + '"]');
            const renameInput = document.querySelector('input[name="rename_' + col + '"]');
            const mapRadio = document.getElementById('map_' + col);
            
            // Function to update form state
            function updateFormState() {
                if (createRadio && typeSelect) {
                    typeSelect.disabled = !createRadio.checked;
                }
                if (renameInput) {
                    renameInput.disabled = !createRadio.checked;
                }
                if (mapOtherRadio && mapOtherSelect) {
                    mapOtherSelect.disabled = !mapOtherRadio.checked;
                }
            }
            
            // Add event listeners
            if (createRadio) {
                createRadio.addEventListener('change', updateFormState);
            }
            if (mapOtherRadio) {
                mapOtherRadio.addEventListener('change', updateFormState);
            }
            if (mapRadio) {
                mapRadio.addEventListener('change', updateFormState);
            }
            
            // Initialize state
            updateFormState();
        });
    });
    </script>
    ''', table=table, df=df, all_table_columns=all_table_columns, 
         existing_column_types=existing_column_types, message=message, 
         type_choices=type_choices, inferred_types=inferred_types) 

@app.route('/view_db', methods=['GET'])
def view_db():
    db_file = session.get('db_choice')
    if not db_file or not os.path.exists(db_file):
        return render_template_string(STYLE + '''
        <div class="header"><img src="/static/Unknown.png" alt="LAB14 Logo" class="header-logo"><span class="header-title">Excel to SQLite Importer</span><nav class="menu"><a href="/" class="menu-link">Start Over</a><a href="/" class="menu-link">View Database</a></nav></div>
        <div class="container">
        <h1>No database selected or file not found.</h1>
        <p>Please go back to the <a href="/">index page</a> to select a database.</p>
        </div>
        ''')
    
    conn = sqlite3.connect(db_file)
    cur = conn.cursor()
    
    # Get all table names
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [row[0] for row in cur.fetchall()]
    
    # Get column metadata for each table
    table_metadata = {}
    for table_name in tables:
        table_metadata[table_name] = get_column_metadata(conn, table_name)
    
    conn.close()
    
    return render_template_string(STYLE + '''
    <div class="header"><img src="/static/Unknown.png" alt="LAB14 Logo" class="header-logo"><span class="header-title">Excel to SQLite Importer</span><nav class="menu"><a href="/" class="menu-link">Start Over</a><a href="/view_db" class="menu-link">View Database</a></nav></div>
    <div class="container">
    <h1>Database: {{db_file}}</h1>
    <div class="export-buttons">
        <a href="{{ url_for('export_db', db_file=db_file) }}">Export Entire Database</a>
    </div>
    <h2>Tables</h2>
    <table>
        <thead>
            <tr>
                <th class="sortable">Table Name</th>
                <th class="sortable">Columns</th>
                <th class="sortable">Actions</th>
            </tr>
        </thead>
        <tbody>
            {% for table_name, metadata in table_metadata.items() %}
            <tr>
                <td>{{ table_name }}</td>
                <td>
                    <ul>
                        {% for col, info in metadata.items() %}
                        <li>
                            {{ col }} ({{ info['type'] }}) - {% if info['is_currency'] %}Currency{% endif %} - {% if info['is_date'] %}Date{% endif %}
                            <a href="{{ url_for('edit_column', db_file=db_file, table_name=table_name, column_name=col) }}" style="margin-left: 10px; font-size: 0.8em; color: #007acc;">[Edit]</a>
                        </li>
                        {% endfor %}
                    </ul>
                </td>
                <td>
                    <a href="{{ url_for('view_table', db_file=db_file, table_name=table_name) }}">View Table</a>
                    <a href="{{ url_for('export_table', db_file=db_file, table_name=table_name) }}">Export Table</a>
                    <a href="{{ url_for('delete_table', db_file=db_file, table_name=table_name) }}">Delete Table</a>
                </td>
            </tr>
            {% endfor %}
        </tbody>
    </table>
    </div>
    ''', db_file=db_file, table_metadata=table_metadata)

@app.route('/edit_column/<path:db_file>/<table_name>/<column_name>', methods=['GET', 'POST'])
def edit_column(db_file, table_name, column_name):
    if not os.path.exists(db_file):
        return "Database file not found.", 404
    
    conn = sqlite3.connect(db_file)
    cur = conn.cursor()
    
    if request.method == 'POST':
        new_column_name = request.form.get('new_column_name', '').strip()
        if not new_column_name:
            message = "Column name cannot be empty."
        elif new_column_name == column_name:
            message = "Column name unchanged."
        else:
            try:
                # Check if new column name already exists
                cur.execute(f"PRAGMA table_info('{table_name}')")
                existing_columns = [row[1] for row in cur.fetchall()]
                
                if new_column_name in existing_columns:
                    message = f"Column '{new_column_name}' already exists in table '{table_name}'."
                else:
                    # Get column metadata before renaming
                    metadata = get_column_metadata(conn, table_name)
                    column_info = metadata.get(column_name, {})
                    
                    # Rename the column using SQLite's ALTER TABLE
                    cur.execute(f'ALTER TABLE "{table_name}" RENAME COLUMN "{column_name}" TO "{new_column_name}"')
                    
                    # Update metadata table if it exists
                    try:
                        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='column_metadata'")
                        if cur.fetchone():
                            cur.execute("UPDATE column_metadata SET column_name = ? WHERE table_name = ? AND column_name = ?", 
                                      (new_column_name, table_name, column_name))
                    except Exception:
                        pass  # Metadata table might not exist
                    
                    conn.commit()
                    message = f"Column '{column_name}' successfully renamed to '{new_column_name}'."
                    
                    # Redirect back to view_db after successful rename
                    conn.close()
                    return redirect(url_for('view_db'))
                    
            except Exception as e:
                message = f"Error renaming column: {e}"
                conn.rollback()
    else:
        message = ""
    
    # Get current column info
    cur.execute(f"PRAGMA table_info('{table_name}')")
    columns = cur.fetchall()
    column_info = None
    for col in columns:
        if col[1] == column_name:
            column_info = {
                'name': col[1],
                'type': col[2],
                'notnull': col[3],
                'default': col[4],
                'pk': col[5]
            }
            break
    
    if not column_info:
        conn.close()
        return f"Column '{column_name}' not found in table '{table_name}'.", 404
    
    # Get metadata for display
    metadata = get_column_metadata(conn, table_name)
    column_metadata = metadata.get(column_name, {})
    
    conn.close()
    
    return render_template_string(STYLE + '''
    <div class="header"><img src="/static/Unknown.png" alt="LAB14 Logo" class="header-logo"><span class="header-title">Excel to SQLite Importer</span><nav class="menu"><a href="/" class="menu-link">Start Over</a><a href="/view_db" class="menu-link">View Database</a></nav></div>
    <div class="container">
    <h1>Edit Column: {{column_name}}</h1>
    <p><strong>Table:</strong> {{table_name}}</p>
    <p><strong>Current Type:</strong> {{column_info.type}}</p>
    {% if column_metadata.is_currency %}
    <p><strong>Special Type:</strong> Currency (stored as INTEGER, displayed as $X.XX)</p>
    {% elif column_metadata.is_date %}
    <p><strong>Special Type:</strong> Date (stored as TEXT, displayed as MM/DD/YYYY)</p>
    {% endif %}
    
    <form method="post">
        <div class="form-group">
            <label for="new_column_name"><strong>New Column Name:</strong></label><br>
            <input type="text" id="new_column_name" name="new_column_name" value="{{column_name}}" style="width: 300px; padding: 8px; margin: 5px 0;">
        </div>
        
        <div class="form-group">
            <input type="submit" value="Rename Column" style="background-color: #007acc; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">
            <a href="/view_db" style="margin-left: 10px; padding: 10px 20px; background-color: #666; color: white; text-decoration: none; border-radius: 4px;">Cancel</a>
        </div>
    </form>
    
    {% if message %}
    <div class="message" style="margin-top: 20px; padding: 10px; background-color: {% if 'Error' in message or 'already exists' in message %}#ffebee{% else %}#e8f5e8{% endif %}; border-radius: 4px;">
        {{message}}
    </div>
    {% endif %}
    </div>
    ''', db_file=db_file, table_name=table_name, column_name=column_name, 
         column_info=column_info, column_metadata=column_metadata, message=message)

@app.route('/export_db/<path:db_file>', methods=['GET'])
def export_db(db_file):
    # URL decode the database file path
    import urllib.parse
    db_file = urllib.parse.unquote(db_file)
    
    if not os.path.exists(db_file):
        return "Database file not found.", 404
    
    try:
        # Check if file is readable
        if not os.access(db_file, os.R_OK):
            return "Database file is not readable.", 403
        
        # Get file size to check if it's reasonable
        file_size = os.path.getsize(db_file)
        if file_size > 100 * 1024 * 1024:  # 100MB limit
            return "Database file is too large to export.", 413
        
        with open(db_file, 'rb') as f:
            return send_file(
                f, 
                as_attachment=True, 
                download_name=os.path.basename(db_file),
                mimetype='application/octet-stream'
            )
    except PermissionError:
        return "Permission denied accessing database file.", 403
    except Exception as e:
        return f"Error exporting database: {str(e)}", 500

@app.route('/export_table/<path:db_file>/<table_name>', methods=['GET'])
def export_table(db_file, table_name):
    if not os.path.exists(db_file):
        return "Database file not found.", 404
    
    try:
        conn = sqlite3.connect(db_file)
        metadata = get_column_metadata(conn, table_name)
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        conn.close()
        
        # Convert currency and date columns for export
        for col in df.columns:
            if col in metadata and metadata[col]['is_currency']:
                df[col] = df[col] / 100
                df[col] = df[col].round(2)
            elif col in metadata and metadata[col]['is_date']:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%m/%d/%Y')
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            workbook = writer.book
            # Get the worksheet name - use a safe default name
            worksheet_name = 'Sheet1'
            worksheet = writer.sheets[worksheet_name]
            # Apply currency format to currency columns
            currency_fmt = None
            try:
                from openpyxl.styles import numbers
                currency_fmt = numbers.FORMAT_CURRENCY_USD_SIMPLE
            except ImportError:
                currency_fmt = '$#,##0.00'
            for idx, col in enumerate(df.columns):
                if col in metadata and metadata[col]['is_currency']:
                    col_letter = chr(ord('A') + idx)
                    worksheet.column_dimensions[col_letter].number_format = currency_fmt
                    for row in range(2, len(df) + 2):
                        cell = worksheet[f'{col_letter}{row}']
                        cell.number_format = currency_fmt
        output.seek(0)
        
        return send_file(output, as_attachment=True, download_name=f"{table_name}.xlsx")
    except Exception as e:
        return f"Error exporting table {table_name}: {e}", 500

@app.route('/delete_table/<path:db_file>/<table_name>', methods=['POST'])
def delete_table(db_file, table_name):
    if not os.path.exists(db_file):
        return "Database file not found.", 404
    
    try:
        conn = sqlite3.connect(db_file)
        cur = conn.cursor()
        cur.execute(f"DROP TABLE IF EXISTS {table_name}")
        conn.commit()
        conn.close()
        return redirect(url_for('view_db'))
    except Exception as e:
        return f"Error deleting table {table_name}: {e}", 500

@app.route('/view_table/<path:db_file>/<table_name>', methods=['GET'])
def view_table(db_file, table_name):
    if not os.path.exists(db_file):
        return "Database file not found.", 404
    
    try:
        conn = sqlite3.connect(db_file)
        
        # Get column metadata
        metadata = get_column_metadata(conn, table_name)
        
        # Get data with optional sorting
        sort_cols = request.args.getlist('sort')
        sort_orders = request.args.getlist('order')
        order_clause = ""
        if sort_cols:
            order_parts = []
            for i, col in enumerate(sort_cols):
                direction = 'ASC'
                if i < len(sort_orders) and sort_orders[i].upper() == 'DESC':
                    direction = 'DESC'
                order_parts.append(f'"{col}" {direction}')
            order_clause = " ORDER BY " + ", ".join(order_parts)
        
        # Get data
        df = pd.read_sql_query(f"SELECT * FROM {table_name}{order_clause}", conn)
        conn.close()
        
        # Debug: Check for problematic column names
        print(f"DEBUG: Table {table_name} columns:")
        for i, col in enumerate(df.columns):
            print(f"  {i}: '{col}' (length: {len(col) if col else 0})")
            if not col or col.strip() == '':
                print(f"    WARNING: Empty or whitespace-only column name at index {i}")
        
        # Convert currency and date columns for display (with proper formatting for DataTables)
        for col in df.columns:
            if col in metadata and metadata[col]['is_currency']:
                # Convert from pennies/cents back to dollars and format
                # First ensure the column is numeric
                df[col] = pd.to_numeric(df[col], errors='coerce')
                df[col] = df[col] / 100
                df[col] = df[col].round(2)
                # Format as currency for display but keep numeric for sorting
                df[col] = df[col].apply(lambda x: f"${x:.2f}" if pd.notna(x) and x != 0 else "")
            elif col in metadata and metadata[col]['is_date']:
                # Convert from ISO format to MM/DD/YYYY for display
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%m/%d/%Y')

        # Limit rows for display
        show_all = request.args.get('show_all') == 'true'
        if not show_all and len(df) > 20:
            df_display = df.head(20)
            has_more = True
        else:
            df_display = df
            has_more = False
        
        return render_template_string(STYLE + '''
        <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.7/css/jquery.dataTables.css">
        <script type="text/javascript" charset="utf8" src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
        <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/1.13.7/js/jquery.dataTables.js"></script>
        <div class="header"><img src="/static/Unknown.png" alt="LAB14 Logo" class="header-logo"><span class="header-title">Excel to SQLite Importer</span><nav class="menu"><a href="/" class="menu-link">Start Over</a><a href="/view_db" class="menu-link">View Database</a></nav></div>
        <div class="container" style="max-width:98vw;">
        <h1>Table: {{table_name}}</h1>
        <div class="export-buttons">
            <a href="{{ url_for('export_table', db_file=db_file, table_name=table_name) }}">Export as Excel</a>
            <a href="{{ url_for('export_table_csv', db_file=db_file, table_name=table_name) }}">Export as CSV</a>
        </div>
        {% if has_more %}
        <p>Showing first 20 rows. <a href="?show_all=true">Show all rows</a></p>
        {% endif %}
        <div style="overflow-x:auto; overflow-y:auto; max-height:70vh; border-radius:8px; box-shadow:0 2px 10px rgba(0,0,0,0.08); background:#fff;">
        <table id="dataTable" style="min-width:1200px; font-size:0.95rem; border-collapse:separate; border-spacing:0;">
            <thead style="position:sticky; top:0; background:#003d21; color:#fff; z-index:2;">
                <tr>
                    {% for col in df_display.columns %}
                    <th style="position:sticky; top:0; background:#003d21; color:#fff; padding:10px 8px; border-bottom:2px solid #ebebeb;">{{col}}</th>
                    {% endfor %}
                </tr>
            </thead>
            <tbody>
                {% for _, row in df_display.iterrows() %}
                <tr>
                    {% for col in df_display.columns %}
                    <td style="padding:8px 8px; border-bottom:1px solid #e7e7e7;">{{row[col]}}</td>
                    {% endfor %}
                </tr>
                {% endfor %}
            </tbody>
        </table>
        </div>
        <script>
        $(document).ready(function() {
            $('#dataTable').DataTable({
                scrollX: true,
                scrollY: '60vh',
                scrollCollapse: true,
                paging: false,
                searching: false,
                info: false,
                ordering: true,
                order: [],
                columnDefs: [
                    {
                        targets: '_all',
                        className: 'dt-body-center',
                        render: function(data, type, row, meta) {
                            if (type === 'display') {
                                return data;
                            }
                            if (type === 'sort') {
                                // Remove $ and convert to number for sorting
                                if (typeof data === 'string' && data.startsWith('$')) {
                                    return parseFloat(data.replace('$', '').replace(',', '')) || 0;
                                }
                                return data;
                            }
                            return data;
                        }
                    }
                ]
            });
        });
        </script>
        ''', table_name=table_name, df_display=df_display, has_more=has_more, db_file=db_file)
        
    except Exception as e:
        return f"Error viewing table {table_name}: {e}", 500

@app.route('/export_table_csv/<path:db_file>/<table_name>', methods=['GET'])
def export_table_csv(db_file, table_name):
    if not os.path.exists(db_file):
        return "Database file not found.", 404
    
    try:
        conn = sqlite3.connect(db_file)
        
        # Get column metadata
        metadata = get_column_metadata(conn, table_name)
        
        # Get data
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        conn.close()
        
        # Convert currency and date columns for export
        for col in df.columns:
            if col in metadata and metadata[col]['is_currency']:
                # Convert from pennies/cents back to dollars
                df[col] = df[col] / 100
                df[col] = df[col].round(2)
            elif col in metadata and metadata[col]['is_date']:
                # Convert from ISO format to MM/DD/YYYY
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%m/%d/%Y')
        
        output = io.StringIO()
        df.to_csv(output, index=False)
        output.seek(0)
        
        return send_file(
            io.BytesIO(output.getvalue().encode('utf-8')),
            as_attachment=True,
            download_name=f"{table_name}.csv",
            mimetype='text/csv'
        )
    except Exception as e:
        return f"Error exporting table {table_name}: {e}", 500

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory('static', filename)

if __name__ == '__main__':
    # Set to '0.0.0.0' for network access, '127.0.0.1' for local only
    app.run(host='127.0.0.1', port=5666, debug=True) 