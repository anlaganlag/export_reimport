import pandas as pd
import os
import time
import argparse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill, Color
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
import shutil

# Make sure outputs directory exists
if not os.path.exists('outputs'):
    try:
        os.makedirs('outputs')
        print("Created outputs directory")
    except Exception as e:
        print(f"Error creating outputs directory: {e}")
        raise

# Function to read Excel files
def read_excel_file(file_path):
    return pd.read_excel(file_path)




def merge_packing_list_cells(workbook_path):
    """
    Merge cells in the Packing List sheet for rows with the same Carton NO.
    Specifically merges the CTNS, Carton MEASUREMENT, G.W (KG), and Carton NO. columns vertically,
    but only for groups with more than one row.
    """
    try:
        wb = load_workbook(workbook_path)
        if 'Packing List' not in wb.sheetnames:
            print("No 'Packing List' sheet found in workbook")
            return False
            
        ws = wb['Packing List']
        
        # Find column indices
        carton_no_idx = None
        ctns_idx = None
        measurement_idx = None
        gw_idx = None
        desc_idx = None
        
        for col_idx, cell in enumerate(ws[1], 1):
            if cell.value == 'Carton NO.':
                carton_no_idx = col_idx
            elif cell.value == 'CTNS':
                ctns_idx = col_idx
            elif cell.value == 'Carton MEASUREMENT':
                measurement_idx = col_idx
            elif cell.value == 'G.W (KG)':
                gw_idx = col_idx
            elif cell.value == 'DESCRIPTION':
                desc_idx = col_idx
        
        if not all([carton_no_idx, ctns_idx, measurement_idx, gw_idx]):
            print("Could not find all required columns for merging")
            return False
            
        # Track rows with the same carton number
        current_carton = None
        start_row = None
        last_row = ws.max_row
        
        # Check if the last row is a "Total" row
        total_row_idx = None
        for row_idx in range(last_row, 1, -1):
            if desc_idx and ws.cell(row=row_idx, column=desc_idx).value == 'Total':
                total_row_idx = row_idx
                break
        
        # If we found a Total row, adjust the last_row to be one before it
        effective_last_row = total_row_idx - 1 if total_row_idx else last_row
        
        for row_idx in range(2, effective_last_row + 1):
            carton_no = ws.cell(row=row_idx, column=carton_no_idx).value
            
            # Skip empty or None values
            if not carton_no:
                continue
                
            # If this is the first carton number encountered
            if current_carton is None:
                current_carton = carton_no
                start_row = row_idx
                continue
                
            # If we encounter a new carton number
            if current_carton != carton_no:
                # Check if the previous group has more than one row
                if start_row < row_idx - 1:  # More than one row in the group
                    end_row = row_idx - 1
                    # Merge CTNS cells
                    ws.merge_cells(start_row=start_row, start_column=ctns_idx, 
                                  end_row=end_row, end_column=ctns_idx)
                    # Merge Carton MEASUREMENT cells
                    ws.merge_cells(start_row=start_row, start_column=measurement_idx,
                                  end_row=end_row, end_column=measurement_idx)
                    # Merge G.W (KG) cells
                    ws.merge_cells(start_row=start_row, start_column=gw_idx,
                                  end_row=end_row, end_column=gw_idx)
                    # Merge Carton NO. cells
                    ws.merge_cells(start_row=start_row, start_column=carton_no_idx,
                                  end_row=end_row, end_column=carton_no_idx)
                
                # Start tracking the new carton
                current_carton = carton_no
                start_row = row_idx
        
        # Handle the last group before the Total row
        if start_row and start_row < effective_last_row:  # More than one row in the last group
            end_row = effective_last_row
            ws.merge_cells(start_row=start_row, start_column=ctns_idx, 
                          end_row=end_row, end_column=ctns_idx)
            ws.merge_cells(start_row=start_row, start_column=measurement_idx,
                          end_row=end_row, end_column=measurement_idx)
            ws.merge_cells(start_row=start_row, start_column=gw_idx,
                          end_row=end_row, end_column=gw_idx)
            ws.merge_cells(start_row=start_row, start_column=carton_no_idx,
                          end_row=end_row, end_column=carton_no_idx)
        
        # Set vertical alignment for merged cells
        for row in ws.iter_rows():
            for cell in row:
                if cell.coordinate in ws.merged_cells:
                    cell.alignment = Alignment(vertical='center')
        
        wb.save(workbook_path)
        print(f"Successfully merged cells in Packing List for {workbook_path}")
        return True
    except Exception as e:
        print(f"Error merging cells in Packing List: {e}")
        return False

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
        '总保费': 10,
        '总运费': 10,
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
            if col_name in ['net weight','FOB单价','Amount', 'FOB总价', 'CIF单价', 'CIF总价(FOB总价+运保费)', '采购单价', '采购总价', '总保费', '总运费', '每公斤摊的运保费', '该项对应的运保费', '单价USD数值']:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif col_name in ['Qty' ]:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right')
            elif col_name in ['Unit Price' ]:
                cell.number_format = '#,##0.000000'
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
    numeric_cols = ['Qty', 'net weight', '采购总价', 'FOB总价', 'CIF总价(FOB总价+运保费)', 'Amount']
    
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
    retry_delay = 2  # seconds
    
    for attempt in range(max_retries):
        try:
            # Safe save to handle file access issues
            df_with_summary.to_excel(file_path, index=False)
            # If we get here, the save was successful
            return True
        except PermissionError:
            if attempt < max_retries - 1:
                print(f"File {file_path} is locked. Retrying in {retry_delay} seconds... (Attempt {attempt+1}/{max_retries})")
                time.sleep(retry_delay)
            else:
                print(f"Could not save to {file_path} after {max_retries} attempts due to permission issues.")
                print("Please close any applications that might have this file open.")
                raise
        except Exception as e:
            print(f"Unexpected error while saving {file_path}: {e}")
            raise

# Function to safely save DataFrame to Excel
def safe_save_to_excel(df, file_path, include_summary=True):
    """Safely save DataFrame to Excel with proper error handling."""
    try:
        if include_summary:
            # Add a summary row
            add_summary_row(df, file_path)
        else:
            # Save directly without summary
            df.to_excel(file_path, index=False)
            
        # Apply styling
        try:
            apply_excel_styling(file_path)
            print(f"Successfully saved and styled: {file_path}")
            return True
        except Exception as e:
            print(f"Warning: File saved but could not apply styling: {e}")
            return True
    except PermissionError as e:
        print(f"Error: Could not save to {file_path} due to permission issues.")
        print(f"Please close the file if it's open in another application.")
        print(f"Error details: {e}")
        return False
    except Exception as e:
        print(f"Error: Failed to save to {file_path}.")
        print(f"Error details: {e}")
        return False

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

    pack_list_expected_columns = [
        'Sr No.',
        'P/N.',
        'DESCRIPTION', 
        'Model NO.',
        'QUANTITY',
        'CTNS',
        'Carton MEASUREMENT',
        'G.W (KG)',
        'N.W(KG)',
        'Carton NO.'
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

def split_by_project_and_factory(df):
    """Split the dataframe by project and factory."""
    print("Available columns for splitting:", df.columns.tolist())
    print("Unique project values:", df['project'].unique())
    print("Unique factory values:", df['factory'].unique())
    
    # Define the project categories with more robust string handling
    project_categories = {
        '大华': lambda x: str(x).strip() == '大华',
        '麦格米特': lambda x: str(x).strip() == '麦格米特',
        '工厂': lambda x: str(x).strip() not in ['大华', '麦格米特']  # Changed from 'others' to '工厂'
    }
    
    # Get unique factories
    factories = df['factory'].unique()
    
    # Dictionary to store split dataframes
    split_dfs = {}
    
    # Split by project and factory
    for project_name, project_filter in project_categories.items():
        try:
            project_df = df[df['project'].apply(project_filter)]
            print(f"Found {len(project_df)} rows for project {project_name}")
            
            for factory in factories:
                key = (project_name, factory)
                split_dfs[key] = project_df[project_df['factory'] == factory]
                print(f"Found {len(split_dfs[key])} rows for {project_name} - {factory}")
        except Exception as e:
            print(f"Error processing project {project_name}: {e}")
            continue
    
    return split_dfs, project_categories

# Main function to process the shipping list
def process_shipping_list(packing_list_file, policy_file, output_dir='outputs'):
    # Read the input files
    packing_list_df = read_excel_file(packing_list_file)
    policy_df = read_excel_file(policy_file)
    
    # Print original column names for debugging
    print("Original packing list columns:")
    for col in packing_list_df.columns:
        print(f"  {col}")
    
    # Define packing list output columns - define this at the beginning so it's available everywhere
    pl_output_columns = [
        'Sr No.', 'P/N.', 'DESCRIPTION', 'Model NO.', 'QUANTITY', 'CTNS', 
        'Carton MEASUREMENT', 'G.W (KG)', 'N.W(KG)', 'Carton NO.'
    ]
    
    print("\nPacking List output columns defined:")
    print(pl_output_columns)
    
    # Extract policy parameters using correct column names
    markup_percentage = policy_df['加价率'].iloc[0]  # Markup percentage
    insurance_coefficient = policy_df['保险系数'].iloc[0]  # Insurance coefficient
    insurance_rate = policy_df['保险费率'].iloc[0]  # Insurance rate
    total_freight_amount = policy_df['总运费(RMB)'].iloc[0]  # Total freight amount
    exchange_rate = policy_df['汇率(RMB/美元)'].iloc[0]  # Exchange rate
    
    # Clean up the column names for better handling
    packing_list_df.columns = [str(col).strip() for col in packing_list_df.columns]
    
    # Create new DataFrames for the processed data
    result_df = pd.DataFrame()
    pl_result_df = pd.DataFrame()
    
    # Keep track of mappings for debugging
    column_mappings = {}
    
    # Find key columns by pattern matching
    print("\nFinding column mappings...")
    # Main invoice columns
    sr_no_col = find_column_with_pattern(packing_list_df, ['sr no', '序号', '序列号'], 'NO.')
    material_code_col = find_column_with_pattern(packing_list_df, ['p/n', '料号', 'material code', '系统料号'], 'Material code')
    description_col = find_column_with_pattern(packing_list_df, ['清关英文货描', '清关英文货描(关务提供)'], 'DESCRIPTION')
    model_col = find_column_with_pattern(packing_list_df, ['model', '型号', '物料型号', '货物型号'], 'Model NO.')
    unit_price_col = find_column_with_pattern(packing_list_df, ['单价', 'unit price', 'price', '不含税单价'], 'Unit Price')
    qty_col = find_column_with_pattern(packing_list_df, ['qty', 'quantity', '数量'], 'Qty')
    unit_col = find_column_with_pattern(packing_list_df, ['unit', '单位', '单位中文'], 'Unit')
    net_weight_col = find_column_with_pattern(packing_list_df, ['net weight', 'N.W  (KG)总净重', 'n.w', '总净重', 'N.W(KG)', 'N.W  (KG)总净重'], 'net weight')
    gross_weight_col = find_column_with_pattern(packing_list_df, ['G.W（KG)\n总毛重','gross weight', 'G.W  (KG)总毛重', 'g.w', '总毛重', 'G.W (KG)', 'G.W  (KG)总毛重'], 'gross weight')
    factory_col = find_column_with_pattern(packing_list_df, ['factory', '工厂', 'daman/silvass'], 'factory')
    project_col = find_column_with_pattern(packing_list_df, ['project', '项目名称', '项目'], 'project')
    end_use_col = find_column_with_pattern(packing_list_df, ['end use', '用途'], 'end use')
    
    # Additional packing list columns
    ctns_col = find_column_with_pattern(packing_list_df, ['ctns', '件数'], 'CTNS')
    # find_column_with_pattern takes 3 parameters:
    # 1. packing_list_df: The dataframe to search in
    # 2. ['体积（CBM）', '总体积', 'measurement']: Array of possible column names to look for in Chinese/English
    #    - Will match any column that contains these strings
    # 3. 'Carton MEASUREMENT': The standardized name to use if a match is found
    #    - This is what the matched column will be renamed to in the output
    carton_measurement_col = find_column_with_pattern(packing_list_df, ['体积（CBM）', '总体积', 'CBM'], 'Carton MEASUREMENT')
    carton_no_col = find_column_with_pattern(packing_list_df, ['carton no', '箱号', 'ctn no'], 'Carton NO.')
    
    # 贸易类型列
    trade_type_col = find_column_with_pattern(packing_list_df, ['出口报关方式', '贸易方式', 'trade type'], 'Trade Type')
    
    # Map main invoice columns
    if sr_no_col:
        result_df['NO.'] = packing_list_df[sr_no_col]
        column_mappings['NO.'] = sr_no_col
    
    if material_code_col:
        result_df['Material code'] = packing_list_df[material_code_col]
        column_mappings['Material code'] = material_code_col
    
    if description_col:
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
        # First, copy the original units to both dataframes
        result_df['Unit'] = packing_list_df[unit_col]
        result_df['Original_Unit'] = packing_list_df[unit_col]  # Keep original units for reference
        pl_result_df['Original_Unit'] = packing_list_df[unit_col]  # Also keep in pl_result_df
        column_mappings['Unit'] = unit_col
    
    if net_weight_col:
        result_df['net weight'] = packing_list_df[net_weight_col]
        column_mappings['net weight'] = net_weight_col
    
    if gross_weight_col:
        result_df['gross weight'] = packing_list_df[gross_weight_col]
        result_df['G.W (KG)'] = packing_list_df[gross_weight_col]  # Also add as G.W (KG)
        column_mappings['gross weight'] = gross_weight_col
    
    if factory_col:
        result_df['factory'] = packing_list_df[factory_col]
        column_mappings['factory'] = factory_col
    
    if project_col:
        result_df['project'] = packing_list_df[project_col]
        column_mappings['project'] = project_col
    
    if end_use_col:
        result_df['end use'] = packing_list_df[end_use_col]
        column_mappings['end use'] = end_use_col
    
    # Map packing list columns
    if sr_no_col:
        pl_result_df['Sr No.'] = packing_list_df[sr_no_col]
    
    if material_code_col:
        pl_result_df['P/N.'] = packing_list_df[material_code_col]
    
    if description_col:
        pl_result_df['DESCRIPTION'] = packing_list_df[description_col]
    
    if model_col:
        pl_result_df['Model NO.'] = packing_list_df[model_col]
    
    if qty_col:
        pl_result_df['QUANTITY'] = packing_list_df[qty_col]
    
    if ctns_col:
        pl_result_df['CTNS'] = packing_list_df[ctns_col]
    else:
        pl_result_df['CTNS'] = 1  # Default value if not found
    
    if carton_measurement_col:
        pl_result_df['Carton MEASUREMENT'] = packing_list_df[carton_measurement_col]
    else:
        pl_result_df['Carton MEASUREMENT'] = ""  # Default empty value if not found
    
    if gross_weight_col:
        pl_result_df['G.W (KG)'] = packing_list_df[gross_weight_col]
    else:
        pl_result_df['G.W (KG)'] = ""  # Default empty value if not found
    
    if net_weight_col:
        pl_result_df['N.W(KG)'] = packing_list_df[net_weight_col]
    else:
        pl_result_df['N.W(KG)'] = 0  # Default value if not found
    
    if carton_no_col:
        pl_result_df['Carton NO.'] = packing_list_df[carton_no_col]
    else:
        pl_result_df['Carton NO.'] = ""  # Default empty value if not found
    
    # Add Trade Type column
    if trade_type_col:
        result_df['Trade Type'] = packing_list_df[trade_type_col]
        pl_result_df['Trade Type'] = packing_list_df[trade_type_col]
        column_mappings['Trade Type'] = trade_type_col
    else:
        # Try to find 出口报关方式 column
        report_type_col = find_column_with_pattern(packing_list_df, ['出口报关方式'], '出口报关方式')
        if report_type_col:
            result_df['Trade Type'] = packing_list_df[report_type_col]
            pl_result_df['Trade Type'] = packing_list_df[report_type_col]
            column_mappings['Trade Type'] = report_type_col
        else:
            print("WARNING: 无法确定贸易类型，默认将所有物料视为一般贸易处理")
            result_df['Trade Type'] = '一般贸易'  # Default to general trade
            pl_result_df['Trade Type'] = '一般贸易'  # Default to general trade
    
    # Add factory to packing list
    if factory_col:
        pl_result_df['factory'] = packing_list_df[factory_col]
    
    # Map project column directly from source
    if '项目名称' in packing_list_df.columns:
        result_df['project'] = packing_list_df['项目名称']
        pl_result_df['project'] = packing_list_df['项目名称']
        column_mappings['project'] = '项目名称'
        print(f"Successfully mapped project column from '项目名称'")
    else:
        print(f"Warning: Project column '项目名称' not found in packing list")
        result_df['project'] = '工厂'  # Changed from 'others' to '工厂'
        pl_result_df['project'] = '工厂'  # Changed from 'others' to '工厂'
    
    # Print found mappings for debugging
    print_column_mappings(column_mappings)
    
    # Apply trade type determination
    def determine_trade_type(row_type):
        if pd.isna(row_type):
            return '一般贸易'  # Default to general trade if empty
        
        row_type_str = str(row_type).strip().lower()
        if '买单' in row_type_str:
            return '买单贸易'
        else:
            return '一般贸易'
    
    # Apply trade type determination to both DataFrames
    result_df['Trade Type'] = result_df['Trade Type'].apply(determine_trade_type)
    pl_result_df['Trade Type'] = pl_result_df['Trade Type'].apply(determine_trade_type)
    
    # Count items by trade type
    general_trade_count = (result_df['Trade Type'] == '一般贸易').sum()
    purchase_trade_count = (result_df['Trade Type'] == '买单贸易').sum()
    print(f"\n贸易类型统计：")
    print(f"  一般贸易物料数量: {general_trade_count}")
    print(f"  买单贸易物料数量: {purchase_trade_count}")
    
    # Set Shipper information for both DataFrames
    result_df['Shipper'] = result_df['Trade Type'].apply(
        lambda x: '创想(创想-PCT)' if x == '一般贸易' else 'Unicair(UC-PCT)'
    )
    
    pl_result_df['Shipper'] = pl_result_df['Trade Type'].apply(
        lambda x: '创想(创想-PCT)' if x == '一般贸易' else 'Unicair(UC-PCT)'
    )
    
    # If Amount column is missing, set to None
    if 'Amount' not in result_df.columns:
        result_df['Amount'] = None
        
    # Convert columns to numeric for calculations
    result_df['net weight'] = pd.to_numeric(result_df['net weight'], errors='coerce')
    result_df['Qty'] = pd.to_numeric(result_df['Qty'], errors='coerce')
    result_df['Unit Price'] = pd.to_numeric(result_df['Unit Price'], errors='coerce')
    
    # Also convert packing list numeric columns
    pl_result_df['QUANTITY'] = pd.to_numeric(pl_result_df['QUANTITY'], errors='coerce')
    pl_result_df['G.W (KG)'] = pd.to_numeric(pl_result_df['G.W (KG)'], errors='coerce')
    pl_result_df['N.W(KG)'] = pd.to_numeric(pl_result_df['N.W(KG)'], errors='coerce')
    
    # Fill NaN values with 0 for numerical calculations
    result_df['net weight'] = result_df['net weight'].fillna(0)
    result_df['Qty'] = result_df['Qty'].fillna(0)
    result_df['Unit Price'] = result_df['Unit Price'].fillna(0)
    
    pl_result_df['QUANTITY'] = pl_result_df['QUANTITY'].fillna(0)
    pl_result_df['G.W (KG)'] = pl_result_df['G.W (KG)'].fillna(0)
    pl_result_df['N.W(KG)'] = pl_result_df['N.W(KG)'].fillna(0)
    
    # Calculate total net weight
    net_weight = result_df['net weight']
    total_net_weight = result_df['net weight'].sum()
    
    # Calculate total cost (采购总价) for each row and sum
    result_df['采购总价']  = result_df['Unit Price'] * result_df['Qty'] 
    total_amount =  result_df['采购总价'].sum()
    #总价加价就是总价FOB
    totalFOB = total_amount * (1 + markup_percentage)
    print(f"  总价FOB: ¥{totalFOB:.4f}")


    total_insurance = totalFOB * insurance_coefficient * insurance_rate
    result_df['总保费'] = total_insurance
    print(f"  总保费: ¥{total_insurance:.4f}")

    result_df['总运费'] = total_freight_amount
    print(f"  总运费: ¥{total_freight_amount:.4f}")


    result_df['每公斤摊的运保费'] = (result_df['总保费'] + result_df['总运费']) / total_net_weight

    result_df['该项对应的运保费'] = result_df['每公斤摊的运保费'] * result_df['net weight']

    # 总CIF = 总价FOB+总保费+总运费(policy表上传)
    total_CIF = totalFOB*(1+insurance_coefficient*insurance_rate)+total_freight_amount
    print(f"  总CIF: ¥{total_CIF:.4f}")


            # 每公斤净重CIF
    cif_per_kg = total_CIF / total_net_weight
    print(f"  每公斤净重CIF: ¥{cif_per_kg:.8f}")

    # 每行数据净重*每公斤净重CIF = 该行数据CIF价格
    unit_kg_cif = cif_per_kg * net_weight

        # Calculate FOB price for each item
    result_df['采购单价'] = result_df['Unit Price']
    result_df['采购总价'] = result_df['Unit Price'] * result_df['Qty']
    result_df['FOB单价'] = result_df['Unit Price'] * (1 + markup_percentage)
    result_df['FOB总价'] = result_df['FOB单价'] * result_df['Qty']


    result_df['CIF总价(FOB总价+运保费)'] =  result_df['FOB总价']  + result_df['该项对应的运保费']

    result_df['CIF单价'] = result_df['CIF总价(FOB总价+运保费)'] / result_df['Qty'].replace(0, 1)




    # Summary statistics
    print(f"\nSummary statistics:")
    print(f"  Total items: {len(result_df)}")
    print(f"  Total net weight: {total_net_weight:.2f} kg")
    
    # Calculate unit freight rate (per kg)
    unit_freight_rate = total_freight_amount / total_net_weight if total_net_weight > 0 else 0
    print(f"  Unit freight rate: ¥{unit_freight_rate:.2f} per kg")
    print(f"  Markup percentage: {markup_percentage*100:.1f}%")
    print(f"  Exchange rate: ¥{exchange_rate:.4f} per USD")
    

    



    
    # Calculate CIF price for each item
    # result_df['CIF总价(FOB总价+运保费)'] = result_df['FOB总价'] + result_df['该项对应的运保费']
    # result_df['CIF单价'] = result_df['CIF总价(FOB总价+运保费)'] / result_df['Qty'].replace(0, 1)  # Prevent division by zero
    
    # Calculate USD value
    result_df['单价USD数值'] = result_df['CIF单价'] * exchange_rate
    
    # Fill in the unit column if it exists
    result_df['单位'] = result_df['Unit'] if 'Unit' in result_df.columns else ""
    
    # Calculate USD unit price
    result_df['Unit Price'] = (result_df['Unit Price'] * exchange_rate).round(8)

    # Calculate Amount as Unit Price multiplied by Quantity
    result_df['Amount'] = (result_df['Unit Price'] * result_df['Qty'] ).round(8)

    # Define output column sets
    cif_output_columns = [
        'NO.', 'Material code', 'DESCRIPTION', 'Model NO.', 'Unit Price', 'Qty', 'Unit', 'Amount',
        'net weight', '采购单价', '采购总价', 'FOB单价', 'FOB总价', '总保费', '总运费', '每公斤摊的运保费',
        '该项对应的运保费', 'CIF总价(FOB总价+运保费)', 'CIF单价', '单价USD数值', '单位',
        'factory', 'project', 'end use'
    ]
    
    cif_output_columns = cif_output_columns + ['G.W (KG)']

    # Define output column sets for export invoice
    exportReimport_output_columns = [
        'NO.', 'Material code', 'DESCRIPTION', 'Model NO.', 'Unit Price', 'Qty', 'Unit', 'Amount'
    ]

    # Internal columns needed for calculations (including Trade Type and Shipper)
    internal_columns = cif_output_columns + ['Trade Type', 'Shipper', 'Original_Unit']

    # Packing list internal columns
    pl_output_columns.append('project')  # Add project to the output columns but it will be removed before final export

    pl_internal_columns = pl_output_columns + ['Trade Type', 'Shipper', 'factory']
    
    # Ensure all required columns exist
    for col in internal_columns:
        if col not in result_df.columns:
            result_df[col] = None
    
    for col in pl_internal_columns:
        if col not in pl_result_df.columns:
            pl_result_df[col] = None
    
    # Reindex the dataframe to match the required column order for internal processing
    result_df = result_df.reindex(columns=internal_columns)
    pl_result_df = pl_result_df.reindex(columns=pl_internal_columns)

    # Drop rows with no material code or all NaN values
    result_df = result_df.dropna(subset=['Material code'], how='all')
    result_df = result_df.dropna(how='all')
    
    pl_result_df = pl_result_df.dropna(subset=['P/N.'], how='all')
    pl_result_df = pl_result_df.dropna(how='all')
    
    # Apply formatting to numeric columns
    numeric_columns = ['采购单价', '采购总价', 'FOB单价', 'FOB总价', '总保费', '总运费', 
                      '每公斤摊的运保费', '该项对应的运保费', 'CIF总价(FOB总价+运保费)', 
                      'CIF单价', '单价USD数值']
    
    for col in numeric_columns:
        if col in result_df.columns:
            result_df[col] = result_df[col].round(2)
    
    # Generate the intermediate CIF invoice file (CIF原始发票)
    cif_invoice = result_df.copy()
    pl_invoice = pl_result_df.copy()
    
    # Remove Trade Type and Shipper columns before saving
    if 'Trade Type' in cif_invoice.columns:
        cif_invoice = cif_invoice.drop(columns=['Trade Type'])
    if 'Shipper' in cif_invoice.columns:
        cif_invoice = cif_invoice.drop(columns=['Shipper'])
        
    if 'Trade Type' in pl_invoice.columns:
        pl_invoice = pl_invoice.drop(columns=['Trade Type'])
    if 'Shipper' in pl_invoice.columns:
        pl_invoice = pl_invoice.drop(columns=['Shipper'])
    if 'factory' in pl_invoice.columns:
        pl_invoice = pl_invoice.drop(columns=['factory'])
        
    cif_file_path = os.path.join(output_dir, 'cif_original_invoice.xlsx')
    pl_file_path = os.path.join(output_dir, 'pl_original_invoice.xlsx')
    
    # Save CIF invoice
    safe_save_to_excel(cif_invoice, cif_file_path)
    safe_save_to_excel(pl_invoice, pl_file_path)
    
    # 提取一般贸易的物料
    general_trade_df = result_df[result_df['Trade Type'] == '一般贸易'].copy()
    pl_df = pl_result_df[pl_result_df['Trade Type'] == '一般贸易'].copy()

    print(f"\nGeneral trade count in result_df: {len(general_trade_df)}")
    print(f"General trade count in pl_result_df: {len(pl_df)}")
    
    print("\nColumns in pl_result_df:")
    print(pl_result_df.columns.tolist())
    
    print("\nColumns in pl_df:")
    print(pl_df.columns.tolist() if not pl_df.empty else "pl_df is empty")
    
    # 只有在存在一般贸易物料时才生成出口发票文件
    if not general_trade_df.empty:
        # Generate the export invoice with two sheets - packing list and commercial invoice
        # First, create a copy for packing list (Sheet1)
        packing_list = general_trade_df.copy()
        
        # Remove Trade Type, Shipper, and project columns before saving to Excel
        if 'Trade Type' in packing_list.columns:
            packing_list = packing_list.drop(columns=['Trade Type'])
        if 'Shipper' in packing_list.columns:
            packing_list = packing_list.drop(columns=['Shipper'])
        if 'project' in packing_list.columns:
            packing_list = packing_list.drop(columns=['project'])
        
        # Use original Chinese units for export invoice
        export_invoice = general_trade_df[exportReimport_output_columns].copy()
        # Keep original Chinese units from the source
        if 'Original_Unit' in general_trade_df.columns:
            export_invoice['Unit'] = general_trade_df['Original_Unit']
        
        # Group by Material code and other fields, keeping original units
        export_grouped = export_invoice.groupby(['Material code', 'Unit Price'], as_index=False).agg({
            'Qty': 'sum',
            'NO.': 'first',
            'Unit': 'first',  # This will keep the original Chinese unit
            'Model NO.': 'first',
            'DESCRIPTION': 'first',
        })
        
        # Calculate Amount after grouping
        export_grouped['Amount'] = (export_grouped['Unit Price'] * export_grouped['Qty']).round(2)
        
        # Ensure all required columns exist and in correct order
        for col in exportReimport_output_columns:
            if col not in export_grouped.columns:
                export_grouped[col] = None
        
        # Reindex to match the required column order
        export_grouped = export_grouped.reindex(columns=exportReimport_output_columns)
        
        # Sort by NO. to maintain original ordering
        export_grouped = export_grouped.sort_values('NO.')
        
        # Reset the index to generate sequential numbers
        export_grouped = export_grouped.reset_index(drop=True)
        export_grouped['NO.'] = export_grouped.index + 1

        # Save both sheets to the same Excel file
        export_file_path = os.path.join(output_dir, 'export_invoice.xlsx')
        
        # Try to delete existing file to avoid permission issues
        try:
            if os.path.exists(export_file_path):
                os.remove(export_file_path)
                print(f"Removed existing file: {export_file_path}")
                time.sleep(1)  # Give the OS time to fully release the file
        except Exception as e:
            print(f"Warning: Could not remove existing file: {e}")
        
        # Create a new Excel writer
        with pd.ExcelWriter(export_file_path, engine='openpyxl') as writer:
            # Packing List 工作表处理
            if not pl_df.empty:
                packing_df = pl_df.copy()
                
                # 确保所有需要的列都存在
                for col in pl_output_columns:
                    if col not in packing_df.columns:
                        print(f"Warning: Column '{col}' not found in packing list data, adding empty column")
                        packing_df[col] = None
                
                # 移除 project 列
                if 'project' in packing_df.columns:
                    packing_df = packing_df.drop(columns=['project'])
                
                # 确保正确的输出列顺序（不包含 project）
                output_columns = [col for col in pl_output_columns if col != 'project']
                packing_df = packing_df[output_columns]
                
                # 添加汇总行（只对数字列计算总和）
                summary_cols = ['QUANTITY', 'G.W (KG)', 'N.W(KG)']
                summary_packing = {}
                for col in summary_cols:
                    if col in packing_df.columns:
                        # Calculate sum without modifying the original column in place
                        # Coerce to numeric, fill NA with 0 JUST for the sum calculation
                        summary_packing[col] = pd.to_numeric(packing_df[col], errors='coerce').fillna(0).sum()

                summary_row = pd.DataFrame([{col: (summary_packing.get(col, None) if col in summary_cols else None) for col in packing_df.columns}])
                summary_row['DESCRIPTION'] = 'Total'
                packing_df = pd.concat([packing_df, summary_row], ignore_index=True)
                
                # Debug print columns
                print("\nPacking List columns before saving:")
                print(packing_df.columns.tolist())
                
                packing_df.to_excel(writer, sheet_name='Packing List', index=False)
            else:
                # 如果没有pl_df数据，创建一个空的packing list with correct columns (不包含 project)
                empty_pl_df = pd.DataFrame(columns=[col for col in pl_output_columns if col != 'project'])
                empty_pl_df.to_excel(writer, sheet_name='Packing List', index=False)

            # Commercial Invoice 工作表处理
            commercial_df = export_grouped.copy()
            # 添加汇总行
            summary_commercial = commercial_df[['Qty', 'Amount']].sum()
            commercial_df = pd.concat([commercial_df, summary_commercial.to_frame().T], ignore_index=True)
            commercial_df.to_excel(writer, sheet_name='Commercial Invoice', index=False)

            # 确保至少一个工作表可见
            workbook = writer.book
            for sheet in workbook.worksheets:
                sheet.sheet_state = "visible"  # 显式设置所有工作表可见
            
            # 设置默认打开第二个sheet
            worksheet = workbook['Commercial Invoice']
            workbook.active = workbook.index(worksheet)
        
        # Apply styling to both sheets
        try:
            # Load the workbook
            wb = load_workbook(export_file_path)
            
            # Style each sheet
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                
                # Define styles
                header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
                header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
                header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                
                # Apply styling to headers
                for col_idx, col in enumerate(ws[1], 1):
                    col.font = header_font
                    col.fill = header_fill
                    col.alignment = header_alignment
                    
                    # Set column width - customize widths for specific columns
                    col_name = ws.cell(row=1, column=col_idx).value
                    if col_name == 'Material code':
                        ws.column_dimensions[get_column_letter(col_idx)].width = 35  # Wider for Material code
                    elif col_name == 'Unit Price':
                        ws.column_dimensions[get_column_letter(col_idx)].width = 20  # Wider for Unit Price
                    elif col_name == 'DESCRIPTION':
                        ws.column_dimensions[get_column_letter(col_idx)].width = 30  # Wider for Description
                    elif col_name == 'Model NO.':
                        ws.column_dimensions[get_column_letter(col_idx)].width = 20  # Wider for Model NO.
                    else:
                        ws.column_dimensions[get_column_letter(col_idx)].width = 15  # Default width
                
                # Apply borders to all cells
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
            
            # Apply number formatting to specific columns
            unit_price_col = None
            amount_col = None
            
            # Find the Unit Price and Amount column indices
            for col_idx, cell in enumerate(ws[1], 1):
                if cell.value == 'Unit Price':
                    unit_price_col = col_idx
                elif cell.value == 'Amount':
                    amount_col = col_idx
            
            # Apply formatting to the entire column
            if unit_price_col:
                for row in range(2, ws.max_row + 1):  # Start from row 2 (skip header)
                    cell = ws.cell(row=row, column=unit_price_col)
                    cell.number_format = '#,##0.000000'
            
            if amount_col:
                for row in range(2, ws.max_row + 1):  # Start from row 2 (skip header)
                    cell = ws.cell(row=row, column=amount_col)
                    cell.number_format = '#,##0.00'
            
            # Save the styled workbook
            wb.save(export_file_path)
            
            # Apply cell merging for packing list
            merge_packing_list_cells(export_file_path)
            
            print(f"Successfully saved and styled export file with multiple sheets: {export_file_path}")
            
            # After saving the export_invoice.xlsx, now merge it with h.xlsx and f.xlsx
            print("Merging files: h.xlsx, export_invoice.xlsx, f.xlsx")
            
            # Temporarily save export_invoice.xlsx to a backup file
            temp_export_file = os.path.join(output_dir, 'temp_export_invoice.xlsx')
            try:
                import shutil
                shutil.copy(export_file_path, temp_export_file)
                
                # Prepare file paths for merging
                h_file = 'h.xlsx'
                f_file = 'f.xlsx'
                
                # For Packing List sheet, we need different files
                pl_h_file = 'pl_h.xlsx'  # First file for Packing List sheet
                pl_f_file = 'pl_f.xlsx'  # Last file for Packing List sheet
                
                # Function to find a file in multiple locations
                def find_file(filename):
                    # Check in current directory
                    if os.path.exists(filename):
                        return filename
                    
                    # Check in output_dir
                    file_in_output = os.path.join(output_dir, filename)
                    if os.path.exists(file_in_output):
                        return file_in_output
                    
                    # Check in the script's directory
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                    file_in_script_dir = os.path.join(script_dir, filename)
                    if os.path.exists(file_in_script_dir):
                        return file_in_script_dir
                    
                    return None  # File not found
                
                # Find the files
                h_file_path = find_file(h_file)
                f_file_path = find_file(f_file)
                pl_h_file_path = find_file(pl_h_file)
                pl_f_file_path = find_file(pl_f_file)
                
                if h_file_path and f_file_path:
                    print(f"Found files for merging Commercial Invoice:")
                    print(f"  {h_file}: {h_file_path}")
                    print(f"  {f_file}: {f_file_path}")
                    
                    if pl_h_file_path and pl_f_file_path:
                        print(f"Found files for merging Packing List:")
                        print(f"  {pl_h_file}: {pl_h_file_path}")
                        print(f"  {pl_f_file}: {pl_f_file_path}")
                    else:
                        print("Warning: Packing List merge files not found, will only merge Commercial Invoice")
                    
                    # Import functions from merge.py
                    import importlib.util
                    import sys
                    
                    # Get the path to merge.py (checking multiple locations)
                    merge_py_path = find_file('merge.py')
                    
                    if merge_py_path:
                        print(f"Found merge.py at: {merge_py_path}")
                        
                        # Call merge.py with the files in specific order using subprocess
                        import subprocess
                        
                        # Prepare command with or without Packing List files
                        if pl_h_file_path and pl_f_file_path:
                            merge_cmd = [
                                sys.executable, 
                                merge_py_path, 
                                h_file_path, 
                                temp_export_file, 
                                f_file_path, 
                                export_file_path,
                                pl_h_file_path,
                                pl_f_file_path
                            ]
                        else:
                            merge_cmd = [
                                sys.executable, 
                                merge_py_path, 
                                h_file_path, 
                                temp_export_file, 
                                f_file_path, 
                                export_file_path
                            ]
                            
                        print(f"Running merge command: {' '.join(merge_cmd)}")
                        
                        try:
                            # Use subprocess.run with stdout and stderr captured to diagnose issues
                            result = subprocess.run(
                                merge_cmd, 
                                check=True,
                                capture_output=True,
                                text=True
                            )
                            
                            # Print stdout and stderr for debugging
                            if result.stdout:
                                print("Merge output:")
                                print(result.stdout)
                            
                            if result.stderr:
                                print("Merge errors:")
                                print(result.stderr)
                            
                            # If successful, verify the file exists
                            if os.path.exists(export_file_path):
                                print(f"Successfully merged files into: {export_file_path}")
                            else:
                                print(f"Error: Merged file not created at {export_file_path}")
                        except subprocess.CalledProcessError as e:
                            print(f"Error running merge script: {e}")
                            if hasattr(e, 'stderr') and e.stderr:
                                print(f"Error details: {e.stderr}")
                        except Exception as e:
                            print(f"Error during file merging: {e}")
                            # Restore original export_invoice.xlsx if merging failed
                            if os.path.exists(temp_export_file):
                                shutil.copy(temp_export_file, export_file_path)
                                print("Restored original export_invoice.xlsx")
                    else:
                        print(f"Error: merge.py not found")
                else:
                    missing_files = []
                    if not h_file_path:
                        missing_files.append(h_file)
                    if not f_file_path:
                        missing_files.append(f_file)
                    if missing_files:
                        print(f"Warning: Could not merge files. Missing files: {', '.join(missing_files)}")
                    else:
                        print("Unexpected error: All files found but merging could not proceed")
            except Exception as e:
                print(f"Error during file merging: {e}")
            finally:
                # Clean up temporary file
                if os.path.exists(temp_export_file):
                    try:
                        os.remove(temp_export_file)
                    except:
                        pass
        except Exception as e:
            print(f"Warning: File saved but could not apply styling: {e}")
    else:
        print("没有一般贸易的物料，不生成出口发票文件")
    
    # After creating result_df and before generating any output files
    # Split the data by project and factory
    split_dfs, project_categories = split_by_project_and_factory(result_df)
    
    # Generate separate invoice files for each split
    for (project, factory), df in split_dfs.items():
        if not df.empty:
            # Create filename
            filename = f'invoice_{project}_{factory}.xlsx'
            file_path = os.path.join(output_dir, filename)
            
            # Create a copy for the invoice
            invoice_df = df[exportReimport_output_columns].copy()
            
            # Create a copy for the packing list
            packing_df = pl_result_df[pl_result_df['factory'] == factory].copy()
            project_filter = project_categories[project]
            packing_df = packing_df[packing_df['project'].apply(project_filter)]
            
            # Save both sheets to the same Excel file
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Save packing list
                if not packing_df.empty:
                    # Remove internal columns before saving
                    save_columns = [col for col in pl_output_columns if col != 'project']  # Remove project from output
                    packing_df = packing_df[save_columns]
                    packing_df.to_excel(writer, sheet_name='Packing List', index=False)
                else:
                    pd.DataFrame(columns=[col for col in pl_output_columns if col != 'project']).to_excel(writer, sheet_name='Packing List', index=False)
                
                # Save commercial invoice
                invoice_df.to_excel(writer, sheet_name='Commercial Invoice', index=False)
            
            # Apply styling
            try:
                apply_excel_styling(file_path)
                # Apply cell merging for packing list
                merge_packing_list_cells(file_path)
                print(f"Successfully generated invoice for project {project}, factory {factory}")
            except Exception as e:
                print(f"Warning: Could not apply styling to {filename}: {e}")

    return result_df


    add_columns = [
        '件数',  # 件数
        '总体积',  # 总体积
        'G.W（KG)\n总毛重',  # 总毛重
        'N.W  (KG)\n总净重',  # 总净重
        'CTN NO.\n(箱号)'  # 箱号
    ]

# Run the process
if __name__ == "__main__":
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description='处理装运清单并生成出口和复进口发票')
    
    parser.add_argument('--packing-list', type=str, default='testfiles/original_packing_list.xlsx',
                      help='原始装箱单文件路径 (默认: testfiles/original_packing_list.xlsx)')
    
    parser.add_argument('--policy', type=str, default='testfiles/policy.xlsx',
                      help='政策文件路径 (默认: testfiles/policy.xlsx)')
    
    parser.add_argument('--output-dir', type=str, default='outputs',
                      help='输出目录 (默认: outputs)')
    
    args = parser.parse_args()
    
    # Create output directory if it doesn't exist
    if not os.path.exists(args.output_dir):
        try:
            os.makedirs(args.output_dir)
            print(f"Created output directory: {args.output_dir}")
        except Exception as e:
            print(f"Error creating output directory: {e}")
            raise
    
    # Get file paths from arguments
    packing_list_file = args.packing_list
    policy_file = args.policy
    
    try:
        # Verify input files exist
        if not os.path.exists(packing_list_file):
            raise FileNotFoundError(f"原始装箱单文件不存在: {packing_list_file}")
        
        if not os.path.exists(policy_file):
            raise FileNotFoundError(f"政策文件不存在: {policy_file}")
        
        result = process_shipping_list(packing_list_file, policy_file, args.output_dir)
        print(f"处理完成！输出文件已保存到 '{args.output_dir}' 目录。")
    except FileNotFoundError as e:
        print(f"错误: {e}")
    except Exception as e:
        print(f"处理文件时出错: {e}")
        import traceback
        traceback.print_exc()
        
        # 打印列名进行调试
        try:
            policy_df = read_excel_file(policy_file)
            print("可用的政策列名:", list(policy_df.columns))
        except:
            print("无法读取政策文件")
            
        try:
            packing_list_df = read_excel_file(packing_list_file)
            print("可用的装箱单列名:", list(packing_list_df.columns))
        except:
            print("无法读取装箱单文件")