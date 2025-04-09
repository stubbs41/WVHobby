import pandas as pd

def read_excel_file(file_path):
    try:
        print(f"\n===== File: {file_path} =====")
        
        # Try to read the Excel file
        df = pd.read_excel(file_path)
        
        # Print columns
        print('Columns:', df.columns.tolist())
        
        # Print first few rows
        print('\nFirst few rows:')
        print(df.head(3))
        
        # Print data types
        print('\nData Types:')
        print(df.dtypes)
        
        # Print basic stats
        print('\nFile Info:')
        print(f"Rows: {len(df)}")
        print(f"Columns: {len(df.columns)}")
        
        return df
    except Exception as e:
        print(f"Error reading {file_path}: {type(e).__name__}: {e}")
        return None

# Read both Excel files
print("Reading 17633.xlsx...")
file1 = read_excel_file('sample/17633.xlsx')

print("\n" + "="*50 + "\n")

print("Reading PO12345.xlsx...")
file2 = read_excel_file('sample/PO12345.xlsx') 