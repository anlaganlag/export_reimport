import pandas as pd
import os

def print_file_info(file_path, description):
    """Print summary information about an Excel file."""
    if not os.path.exists(file_path):
        print(f"{description} file not found: {file_path}")
        return
    
    try:
        df = pd.read_excel(file_path)
        print(f"\n{description} ({file_path}):")
        print(f"  - Number of rows: {len(df)}")
        print(f"  - Number of columns: {len(df.columns)}")
        print(f"  - Columns: {list(df.columns)}")
        
        # Print summary statistics for key numerical columns
        numeric_columns = ['FOB单价', 'FOB总价', 'CIF单价', 'CIF总价(FOB总价+运保费)']
        for col in numeric_columns:
            if col in df.columns:
                print(f"\n  {col} summary:")
                print(f"    - Sum: {df[col].sum():.2f}")
                print(f"    - Mean: {df[col].mean():.2f}")
                print(f"    - Min: {df[col].min():.2f}")
                print(f"    - Max: {df[col].max():.2f}")
        
        # Print the first 2 rows for a preview
        print("\n  Preview (first 2 rows):")
        if len(df) > 0:
            preview_df = df.head(2)
            for idx, row in preview_df.iterrows():
                print(f"    Row {idx}:")
                for col in ['Material code', 'Unit Price', 'Qty', 'FOB单价', 'FOB总价', 'CIF单价', 'factory']:
                    if col in df.columns:
                        print(f"      {col}: {row[col]}")
    except Exception as e:
        print(f"Error reading {file_path}: {e}")

# Check the export invoice file
export_file = 'outputs/export_invoice.xlsx'
print_file_info(export_file, "Export Invoice")

# Check the reimport invoice files
for filename in os.listdir('outputs'):
    if filename.startswith('reimport_invoice_factory_') and filename.endswith('.xlsx'):
        file_path = os.path.join('outputs', filename)
        factory_name = filename.replace('reimport_invoice_factory_', '').replace('.xlsx', '')
        print_file_info(file_path, f"Reimport Invoice for {factory_name}")

print("\nOutput files check complete.")