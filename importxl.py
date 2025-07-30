import pandas as pd
import sqlite3

# Path to your Excel file
excel_file = '/Users/niels/Python/Funnel14/Data/HIMT US Opportunities-2025-07-10-10-00-02.xlsx'
# Name of the SQLite database file
db_file = 'my_database.db'
# Name of the table to insert data into
table_name = 'my_table'

# Read Excel file (first sheet)
df = pd.read_excel(excel_file, sheet_name=0)

# Connect to SQLite database (it will be created if it doesn't exist)
conn = sqlite3.connect(db_file)
cursor = conn.cursor()

# Get existing columns from the table
def refresh_table_columns():
    cursor.execute(f'PRAGMA table_info({table_name})')
    table_info = cursor.fetchall()
    return [col[1] for col in table_info]

table_columns = refresh_table_columns()

# Check for missing columns
new_df = df.copy()
for col in df.columns:
    if col not in table_columns:
        print(f"Column '{col}' not found in table '{table_name}'.")
        print("Select a column to map to, 0 to skip, or 'n' to create a new column:")
        for idx, tcol in enumerate(table_columns, 1):
            print(f"{idx}: {tcol}")
        while True:
            selection = input(f"Enter number for mapping '{col}', 0 to skip, or 'n' to create new column: ").strip()
            if selection == '0':
                new_df = new_df.drop(columns=[col])
                break
            elif selection.lower() == 'n':
                # Infer SQLite type from pandas dtype
                dtype = new_df[col].dtype
                if pd.api.types.is_integer_dtype(dtype):
                    sqltype = 'INTEGER'
                elif pd.api.types.is_float_dtype(dtype):
                    sqltype = 'REAL'
                elif pd.api.types.is_bool_dtype(dtype):
                    sqltype = 'BOOLEAN'
                else:
                    sqltype = 'TEXT'
                try:
                    cursor.execute(f'ALTER TABLE {table_name} ADD COLUMN "{col}" {sqltype}')
                    conn.commit()
                    print(f"Added column '{col}' of type {sqltype} to table '{table_name}'.")
                    table_columns = refresh_table_columns()
                except Exception as e:
                    print(f"Failed to add column: {e}")
                    new_df = new_df.drop(columns=[col])
                break
            else:
                try:
                    idx = int(selection)
                    if 1 <= idx <= len(table_columns):
                        new_df = new_df.rename(columns={col: table_columns[idx-1]})
                        break
                    else:
                        print("Invalid selection. Try again.")
                except ValueError:
                    print("Please enter a valid number or 'n'.")

# Only keep columns that exist in the table
new_df = new_df[[c for c in new_df.columns if c in table_columns]]

# Write the data to the database
new_df.to_sql(table_name, conn, if_exists='append', index=False)

conn.close()
print(f"Data from {excel_file} added to {table_name} in {db_file}")