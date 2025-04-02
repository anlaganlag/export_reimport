import pandas as pd
import os
import time
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill, Color
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# Ensure the outputs directory exists
if not os.path.exists('outputs'):
    os.makedirs('outputs')

# Function to read Excel files
def read_excel_file(file_path):
    return pd.read_excel(file_path)

# Function to apply styling to Excel workbook
def apply_excel_styling(file_path):
    """Apply professional styling to the Excel workbook."""
    # Load the workbook
    wb = load_workbook(file_path)
    ws = wb.active
    
    # Define styles
    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Column widths
    column_widths = {
        'NO.': 6,
        'Material code': 20,
        'DESCRIPTION': 30,
        'Model NO.': 20,
        'Unit Price': 12,
        'Qty': 8,
        'Unit': 8,
        'Amount': 12,
        'net weight': 12,
        '采购单价': 12,
        '采购总价': 12,
        'FOB单价': 12,
        'FOB总价': 12,
        '保费': 10,
        '运费': 10,
        '每公斤摊的运保费': 18,
        '该项对应的运保费': 18,
        'CIF总价(FOB总价+运保费)': 20,
        'CIF单价': 12,
        '单价USD数值': 12,
        '单位': 8,
        '开票品名': 15,
        'factory': 12,
        'project': 15,
        'end use': 15
    }
    
    # Apply header styling
    for col_idx, col in enumerate(ws[1], 1):
        col.font = header_font
        col.fill = header_fill
        col.alignment = header_alignment
        
        # Set column width if specified
        col_name = ws.cell(row=1, column=col_idx).value
        if col_name in column_widths:
            ws.column_dimensions[get_column_letter(col_idx)].width = column_widths[col_name]
        else:
            ws.column_dimensions[get_column_letter(col_idx)].width = 12
    
    # Apply data cell styling
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
        for col_idx, cell in enumerate(row, 1):
            col_name = ws.cell(row=1, column=col_idx).value
            
            # Apply cell styling based on column type
            if col_name in ['FOB单价', 'FOB总价', 'CIF单价', 'CIF总价(FOB总价+运保费)', '采购单价', '采购总价', '保费', '运费', '每公斤摊的运保费', '该项对应的运保费', '单价USD数值']:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif col_name in ['Qty', 'net weight']:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif col_name in ['NO.']:
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.alignment = Alignment(vertical='center', wrap_text=True)
    
    # Add borders to all cells
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border
    
    # Freeze the header row
    ws.freeze_panes = 'A2'
    
    # Save the styled workbook
    wb.save(file_path)

# Function to add summary row to Excel file
def add_summary_row(df, file_path):
    """Add a summary row to the Excel file and save."""
    # Calculate totals for numeric columns
    numeric_cols = ['Qty', 'net weight', '采购总价', 'FOB总价', '保费', '运费', '该项对应的运保费', 'CIF总价(FOB总价+运保费)']
    
    # Create a summary row
    summary = {}
    summary['Material code'] = 'Total'
    
    for col in numeric_cols:
        if col in df.columns:
            summary[col] = df[col].sum()
    
    # Append summary row to DataFrame
    summary_df = pd.DataFrame([summary])
    df_with_summary = pd.concat([df, summary_df], ignore_index=True)
    
    # Write to Excel with retries to handle permission issues
    max_retries = 3
    for attempt in range(max_retries):
        try:
            # Safe save to handle file access issues
            df_with_summary.to_excel(file_path, index=False)
            break
        except PermissionError:
            if attempt < max_retries - 1:
                print(f"File {file_path} is locked. Retrying in 2 seconds... (Attempt {attempt+1}/{max_retries})")
                time.sleep(2)
            else:
                print(f"Could not save to {file_path} after {max_retries} attempts due to permission issues.")
                raise

# Helper function to find columns with specific patterns
def find_column_with_pattern(df, patterns, target_col_name=None):
    """Find a column that contains any of the given patterns."""
    for col in df.columns:
        col_str = str(col).lower()
        for pattern in patterns:
            if pattern.lower() in col_str:
                if target_col_name:
                    print(f"Found column '{col}' for {target_col_name}")
                return col
    if target_col_name:
        print(f"WARNING: Could not find a column matching patterns {patterns} for {target_col_name}")
    return None

# Helper function to print found mappings
def print_column_mappings(mappings):
    """Print a summary of all column mappings found."""
    print("\nColumn mappings summary:")
    missing_cols = []
    
    # Define the expected column list
    expected_columns = [
        'NO.', 'Material code', 'DESCRIPTION', 'Model NO.', 'Unit Price', 'Qty', 'Unit', 
        'net weight', 'factory', 'project', 'end use'
    ]
    
    # Print the found mappings
    print("\nFound column mappings:")
    for target_col in expected_columns:
        if target_col in mappings:
            print(f"  {target_col} <- {mappings[target_col]}")
        else:
            missing_cols.append(target_col)
    
    # Print missing mappings
    if missing_cols:
        print("\nMissing column mappings:")
        for col in missing_cols:
            print(f"  {col} - No matching column found in the source file")
    
    # Summary
    found_count = len(mappings)
    expected_count = len(expected_columns)
    print(f"\nFound {found_count} out of {expected_count} expected column mappings ({100*found_count/expected_count:.1f}%)")

# Main function to process the shipping list
def process_shipping_list(packing_list_file, policy_file):
    # Read the input files
    packing_list_df = read_excel_file(packing_list_file)
    policy_df = read_excel_file(policy_file)
    
    # Print original column names for debugging
    print("Original packing list columns:")
    for col in packing_list_df.columns:
        print(f"  {col}")
    
    # Extract policy parameters using correct column names
    markup_percentage = policy_df['加价率'].iloc[0]  # Markup percentage
    insurance_coefficient = policy_df['保险系数'].iloc[0]  # Insurance coefficient
    insurance_rate = policy_df['保险费率'].iloc[0]  # Insurance rate
    total_freight_amount = policy_df['总运费(RMB)'].iloc[0]  # Total freight amount
    exchange_rate = policy_df['汇率(RMB/美元)'].iloc[0]  # Exchange rate
    
    # Clean up the column names for better handling
    packing_list_df.columns = [str(col).strip() for col in packing_list_df.columns]
    
    # Create a new DataFrame for the processed data
    result_df = pd.DataFrame()
    
    # Keep track of mappings for debugging
    column_mappings = {}
    
    # Find key columns by pattern matching
    print("\nFinding column mappings...")
    sr_no_col = find_column_with_pattern(packing_list_df, ['sr no', '序号', '序列号'], 'NO.')
    material_code_col = find_column_with_pattern(packing_list_df, ['p/n', '料号', 'material code', '系统料号'], 'Material code')
    
    # Use '开票名称' as the source for DESCRIPTION
    description_col = find_column_with_pattern(packing_list_df, ['开票名称', '开票品名'], 'DESCRIPTION')
    
    model_col = find_column_with_pattern(packing_list_df, ['model', '型号', '物料型号', '货物型号'], 'Model NO.')
    unit_price_col = find_column_with_pattern(packing_list_df, ['单价', 'unit price', 'price', '不含税单价'], 'Unit Price')
    qty_col = find_column_with_pattern(packing_list_df, ['qty', 'quantity', '数量'], 'Qty')
    unit_col = find_column_with_pattern(packing_list_df, ['unit', '单位', '单位中文'], 'Unit')
    net_weight_col = find_column_with_pattern(packing_list_df, ['net weight', '净重', 'n.w', '总净重', '单件净重'], 'net weight')
    factory_col = find_column_with_pattern(packing_list_df, ['factory', '工厂', 'daman/silvass'], 'factory')
    project_col = find_column_with_pattern(packing_list_df, ['project', '项目名称', '项目'], 'project')
    end_use_col = find_column_with_pattern(packing_list_df, ['end use', '用途'], 'end use')
    
    # Map found columns to result DataFrame
    if sr_no_col:
        result_df['NO.'] = packing_list_df[sr_no_col]
        column_mappings['NO.'] = sr_no_col
    
    if material_code_col:
        result_df['Material code'] = packing_list_df[material_code_col]
        column_mappings['Material code'] = material_code_col
    
    # Use '开票名称' for DESCRIPTION directly
    if description_col:
        # Store the 开票名称 values in DESCRIPTION column
        result_df['DESCRIPTION'] = packing_list_df[description_col]
        column_mappings['DESCRIPTION'] = description_col
        print(f"Using '{description_col}' as the source for DESCRIPTION")
    
    if model_col:
        result_df['Model NO.'] = packing_list_df[model_col]
        column_mappings['Model NO.'] = model_col
    
    if unit_price_col:
        result_df['Unit Price'] = packing_list_df[unit_price_col]
        column_mappings['Unit Price'] = unit_price_col
    
    if qty_col:
        result_df['Qty'] = packing_list_df[qty_col]
        column_mappings['Qty'] = qty_col
    
    if unit_col:
        result_df['Unit'] = packing_list_df[unit_col]
        column_mappings['Unit'] = unit_col
    
    if net_weight_col:
        result_df['net weight'] = packing_list_df[net_weight_col]
        column_mappings['net weight'] = net_weight_col
    
    if factory_col:
        result_df['factory'] = packing_list_df[factory_col]
        column_mappings['factory'] = factory_col
    
    if project_col:
        result_df['project'] = packing_list_df[project_col]
        column_mappings['project'] = project_col
    
    if end_use_col:
        result_df['end use'] = packing_list_df[end_use_col]
        column_mappings['end use'] = end_use_col
    
    # Print found mappings for debugging
    print_column_mappings(column_mappings)
    
    # If Amount column is missing, set to None
    if 'Amount' not in result_df.columns:
        result_df['Amount'] = None
    
    # Calculate total net weight from packing list
    # Convert to numeric, handling errors by coercing to NaN
    result_df['net weight'] = pd.to_numeric(result_df['net weight'], errors='coerce')
    result_df['Qty'] = pd.to_numeric(result_df['Qty'], errors='coerce')
    result_df['Unit Price'] = pd.to_numeric(result_df['Unit Price'], errors='coerce')
    
    # Fill NaN values with 0 for numerical calculations
    result_df['net weight'] = result_df['net weight'].fillna(0)
    result_df['Qty'] = result_df['Qty'].fillna(0)
    result_df['Unit Price'] = result_df['Unit Price'].fillna(0)
    
    total_net_weight = result_df['net weight'].sum()
    
    # Summary statistics
    print(f"\nSummary statistics:")
    print(f"  Total items: {len(result_df)}")
    print(f"  Total net weight: {total_net_weight:.2f} kg")
    
    # Calculate unit freight rate (per kg)
    unit_freight_rate = total_freight_amount / total_net_weight if total_net_weight > 0 else 0
    print(f"  Unit freight rate: ¥{unit_freight_rate:.2f} per kg")
    print(f"  Markup percentage: {markup_percentage*100:.1f}%")
    print(f"  Exchange rate: ¥{exchange_rate:.4f} per USD")
    
    # Calculate FOB price for each item
    result_df['采购单价'] = result_df['Unit Price']
    result_df['采购总价'] = result_df['Unit Price'] * result_df['Qty']
    result_df['FOB单价'] = result_df['Unit Price'] * (1 + markup_percentage)
    result_df['FOB总价'] = result_df['FOB单价'] * result_df['Qty']
    
    # Calculate insurance fee for each item
    result_df['保费'] = result_df['FOB单价'] * insurance_coefficient * insurance_rate
    
    # Calculate freight for each item
    result_df['每公斤摊的运保费'] = unit_freight_rate
    result_df['运费'] = result_df['net weight'] * unit_freight_rate
    
    # Calculate total insurance and freight cost for each item
    result_df['该项对应的运保费'] = result_df['运费'] + (result_df['保费'] * result_df['Qty'])
    
    # Calculate CIF price for each item
    result_df['CIF总价(FOB总价+运保费)'] = result_df['FOB总价'] + result_df['该项对应的运保费']
    result_df['CIF单价'] = result_df['CIF总价(FOB总价+运保费)'] / result_df['Qty'].replace(0, 1)  # Prevent division by zero
    
    # Calculate USD value
    result_df['单价USD数值'] = result_df['CIF单价'] / exchange_rate
    
    # Fill in the unit column if it exists
    result_df['单位'] = result_df['Unit'] if 'Unit' in result_df.columns else ""
    
    # Calculate USD unit price
    result_df['Unit Price'] = result_df['Unit Price'] * exchange_rate

    # Calculate Amount as Unit Price multiplied by Quantity
    result_df['Amount'] = result_df['Unit Price'] * result_df['Qty'] 

    # Ensure Amount is included in the output columns
    output_columns = [
        'NO.', 'Material code', 'DESCRIPTION', 'Model NO.', 'Unit Price', 'Qty', 'Unit', 'Amount',
        'net weight', '采购单价', '采购总价', 'FOB单价', 'FOB总价', '保费', '运费', '每公斤摊的运保费',
        '该项对应的运保费', 'CIF总价(FOB总价+运保费)', 'CIF单价', '单价USD数值', '单位',
        'factory', 'project', 'end use'
    ]
    
    # Ensure all required columns exist
    for col in output_columns:
        if col not in result_df.columns:
            result_df[col] = None
    
    # Reindex the dataframe to match the required column order
    result_df = result_df.reindex(columns=output_columns)
    
    # Drop rows with no material code or all NaN values
    result_df = result_df.dropna(subset=['Material code'], how='all')
    result_df = result_df.dropna(how='all')
    
    # Apply formatting to numeric columns
    numeric_columns = ['采购单价', '采购总价', 'FOB单价', 'FOB总价', '保费', '运费', 
                      '每公斤摊的运保费', '该项对应的运保费', 'CIF总价(FOB总价+运保费)', 
                      'CIF单价', '单价USD数值']
    
    for col in numeric_columns:
        if col in result_df.columns:
            result_df[col] = result_df[col].round(2)
    
    # Generate the export invoice with styling
    export_invoice = result_df.copy()
    
    # Adjust column names for the output - renaming to 开票品名 to match Chinese context
    # and meet user requirements
    if '开票品名' in export_invoice.columns and 'DESCRIPTION' in export_invoice.columns:
        # If both columns exist, drop one to avoid duplication
        export_invoice.drop(columns=['开票品名'], inplace=True)
        export_invoice.rename(columns={'DESCRIPTION': '开票品名'}, inplace=True)
    elif 'DESCRIPTION' in export_invoice.columns:
        # If only DESCRIPTION exists, rename it
        export_invoice.rename(columns={'DESCRIPTION': '开票品名'}, inplace=True)
    
    export_file_path = 'outputs/export_invoice.xlsx'
    
    # Add summary and save export invoice
    add_summary_row(export_invoice, export_file_path)
    
    # Apply styling to the export invoice
    try:
        apply_excel_styling(export_file_path)
    except Exception as e:
        print(f"Warning: Could not apply styling to export invoice: {e}")
    
    # Generate reimport invoices by factory with styling
    factories = result_df['factory'].dropna().unique()
    for factory in factories:
        factory_df = result_df[result_df['factory'] == factory].copy()
        if not factory_df.empty:
            factory_file_path = f'outputs/reimport_invoice_factory_{factory}.xlsx'
            
            # Add summary and save factory invoice
            add_summary_row(factory_df, factory_file_path)
            
            # Apply styling to the factory invoice
            try:
                apply_excel_styling(factory_file_path)
            except Exception as e:
                print(f"Warning: Could not apply styling to {factory} invoice: {e}")
    
    return result_df

# Run the process
if __name__ == "__main__":
    packing_list_file = 'testfiles/original_packing_list.xlsx'
    policy_file = 'testfiles/policy.xlsx'
    
    try:
        result = process_shipping_list(packing_list_file, policy_file)
        print(f"Processing complete. Output files saved to the 'outputs' directory.")
    except Exception as e:
        print(f"Error processing files: {e}")
        import traceback
        traceback.print_exc()
        # Print the column names for debugging
        policy_df = read_excel_file(policy_file)
        print("Available policy columns:", list(policy_df.columns))
        packing_list_df = read_excel_file(packing_list_file)
        print("Available packing list columns:", list(packing_list_df.columns))