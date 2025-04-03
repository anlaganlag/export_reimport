import pandas as pd
import os
import time
import argparse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill, Color
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

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
            if col_name in ['FOB单价', 'FOB总价', 'CIF单价', 'CIF总价(FOB总价+运保费)', '采购单价', '采购总价', '总保费', '总运费', '每公斤摊的运保费', '该项对应的运保费', '单价USD数值']:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif col_name in ['Qty', 'net weight']:
                cell.number_format = '#,##0'
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
def process_shipping_list(packing_list_file, policy_file, output_dir='outputs'):
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
    net_weight_col = find_column_with_pattern(packing_list_df, ['net weight', 'N.W  (KG)总净重', 'n.w', '总净重', 'N.W  (KG)总净重'], 'net weight')
    factory_col = find_column_with_pattern(packing_list_df, ['factory', '工厂', 'daman/silvass'], 'factory')
    project_col = find_column_with_pattern(packing_list_df, ['project', '项目名称', '项目'], 'project')
    end_use_col = find_column_with_pattern(packing_list_df, ['end use', '用途'], 'end use')
    
    # 查找贸易类型列
    trade_type_col = find_column_with_pattern(packing_list_df, ['出口报关方式', '贸易方式', 'trade type'], 'Trade Type')
    
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
    
    # 添加贸易类型列
    if trade_type_col:
        result_df['Trade Type'] = packing_list_df[trade_type_col]
        column_mappings['Trade Type'] = trade_type_col
    else:
        # 如果找不到贸易类型列，尝试分析出口报关方式列
        report_type_col = find_column_with_pattern(packing_list_df, ['出口报关方式'], '出口报关方式')
        if report_type_col:
            result_df['Trade Type'] = packing_list_df[report_type_col]
            column_mappings['Trade Type'] = report_type_col
        else:
            print("WARNING: 无法确定贸易类型，默认将所有物料视为一般贸易处理")
            result_df['Trade Type'] = '一般贸易'  # 默认为一般贸易
    
    # Print found mappings for debugging
    print_column_mappings(column_mappings)
    
    # 检查贸易类型
    # 确定每行的贸易类型（一般贸易或买单）
    def determine_trade_type(row_type):
        if pd.isna(row_type):
            return '一般贸易'  # 如果为空，默认为一般贸易
        
        row_type_str = str(row_type).strip().lower()
        if '买单' in row_type_str :
            return '买单贸易'
        else:
            return '一般贸易'
    
    # 应用贸易类型判断
    result_df['Trade Type'] = result_df['Trade Type'].apply(determine_trade_type)
    
    # 统计两种贸易类型的数量
    general_trade_count = (result_df['Trade Type'] == '一般贸易').sum()
    purchase_trade_count = (result_df['Trade Type'] == '买单贸易').sum()
    print(f"\n贸易类型统计：")
    print(f"  一般贸易物料数量: {general_trade_count}")
    print(f"  买单贸易物料数量: {purchase_trade_count}")
    
    # 设置发货人信息 Shipper
    result_df['Shipper'] = result_df['Trade Type'].apply(
        lambda x: '创想(创想-PCT)' if x == '一般贸易' else 'Unicair(UC-PCT)'
    )
    
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
    
    # Calculate total net weightf
    net_weight = result_df['net weight']
    total_net_weight = result_df['net weight'].sum()




    
    # Calculate total cost (采购总价) for each row and sum
    result_df['采购总价']  = result_df['Unit Price'] * result_df['Qty'] 
    total_amount =  result_df['采购总价'].sum()
    #总价加价就是总价FOB
    totalFOB = total_amount * (1 + markup_percentage)
    print(f"  总价FOB: ¥{totalFOB:.4f}")


    total_insurance = total_amount * insurance_coefficient * insurance_rate
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

    result_df['CIF总价(FOB总价+运保费)'] =  (result_df['采购总价']*markup_percentage) * (1+insurance_coefficient*insurance_rate) + unit_kg_cif*net_weight

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
    
    # Calculate FOB price for each item
    result_df['采购单价'] = result_df['Unit Price']
    result_df['采购总价'] = result_df['Unit Price'] * result_df['Qty']
    result_df['FOB单价'] = result_df['Unit Price'] * (1 + markup_percentage)
    result_df['FOB总价'] = result_df['FOB单价'] * result_df['Qty']
    



    
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

    # Ensure Amount is included in the output columns
    cif_output_columns = [
        'NO.', 'Material code', 'DESCRIPTION', 'Model NO.', 'Unit Price', 'Qty', 'Unit', 'Amount',
        'net weight', '采购单价', '采购总价', 'FOB单价', 'FOB总价', '总保费', '总运费', '每公斤摊的运保费',
        '该项对应的运保费', 'CIF总价(FOB总价+运保费)', 'CIF单价', '单价USD数值', '单位',
        'factory', 'project', 'end use'
    ]

    # Ensure Amount is included in the output columns
    exportReimport_output_columns = [
        'NO.', 'Material code', 'DESCRIPTION', 'Model NO.', 'Unit Price', 'Qty', 'Unit', 'Amount'
    ]
    
    # 内部计算需要的完整列集合（包含Trade Type和Shipper）
    internal_columns = cif_output_columns + ['Trade Type', 'Shipper']
    
    # Ensure all required columns exist
    for col in internal_columns:
        if col not in result_df.columns:
            result_df[col] = None
    
    # Reindex the dataframe to match the required column order for internal processing
    result_df = result_df.reindex(columns=internal_columns)
    
    # Drop rows with no material code or all NaN values
    result_df = result_df.dropna(subset=['Material code'], how='all')
    result_df = result_df.dropna(how='all')
    
    # Apply formatting to numeric columns
    numeric_columns = ['采购单价', '采购总价', 'FOB单价', 'FOB总价', '总保费', '总运费', 
                      '每公斤摊的运保费', '该项对应的运保费', 'CIF总价(FOB总价+运保费)', 
                      'CIF单价', '单价USD数值']
    
    for col in numeric_columns:
        if col in result_df.columns:
            result_df[col] = result_df[col].round(2)
    
    # Generate the intermediate CIF invoice file (CIF原始发票)
    cif_invoice = result_df.copy()
    
    # Remove Trade Type and Shipper columns before saving
    if 'Trade Type' in cif_invoice.columns:
        cif_invoice = cif_invoice.drop(columns=['Trade Type'])
    if 'Shipper' in cif_invoice.columns:
        cif_invoice = cif_invoice.drop(columns=['Shipper'])
        
    cif_file_path = os.path.join(output_dir, 'cif_original_invoice.xlsx')
    
    # Save CIF invoice
    safe_save_to_excel(cif_invoice, cif_file_path)
    
    # 提取一般贸易的物料
    general_trade_df = result_df[result_df['Trade Type'] == '一般贸易'].copy()
    
    # 只有在存在一般贸易物料时才生成出口发票文件
    if not general_trade_df.empty:
        # Generate the export invoice with two sheets - packing list and commercial invoice
        # First, create a copy for packing list (Sheet1)
        packing_list = general_trade_df.copy()
        
        # Remove Trade Type and Shipper columns before saving to Excel
        if 'Trade Type' in packing_list.columns:
            packing_list = packing_list.drop(columns=['Trade Type'])
        if 'Shipper' in packing_list.columns:
            packing_list = packing_list.drop(columns=['Shipper'])
        
        # Create a copy for commercial invoice (Sheet2) - with only the required columns
        export_invoice = general_trade_df[exportReimport_output_columns].copy()
        
        # Group by Material code, DESCRIPTION, Model NO., Unit Price, and Unit to merge entries
        # This combines items with the same material code and price
        export_grouped = export_invoice.groupby(['Material code', 'Unit Price', ], as_index=False).agg({
            'Qty': 'sum',
            'NO.': 'first' # Keep the first item number
        })
        
        # Recalculate Amount based on grouped quantities
        export_grouped['Amount'] = export_grouped['Unit Price'] * export_grouped['Qty']
        
        # Ensure all required columns exist
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
        
        # Create a new Excel writer
        with pd.ExcelWriter(export_file_path, engine='openpyxl') as writer:
            # Write the packing list to Sheet1
            packing_list.to_excel(writer, sheet_name='Packing List', index=False)
            
            # Write the commercial invoice to Sheet2
            export_grouped.to_excel(writer, sheet_name='Commercial Invoice', index=False)
        
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
                    
                    # Set column width
                    ws.column_dimensions[get_column_letter(col_idx)].width = 15
                
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
            
            # Save the styled workbook
            wb.save(export_file_path)
            print(f"Successfully saved and styled export file with multiple sheets: {export_file_path}")
        except Exception as e:
            print(f"Warning: File saved but could not apply styling: {e}")
    else:
        print("没有一般贸易的物料，不生成出口发票文件")
    
    # Generate reimport invoices by factory - using only required columns, no grouping needed
    factories = result_df['factory'].dropna().unique()
    for factory in factories:
        factory_df = result_df[result_df['factory'] == factory].copy()
        if not factory_df.empty:
            # Select only required columns for reimport invoice
            factory_df = factory_df[exportReimport_output_columns].copy()
            
            # Ensure Trade Type and Shipper are not included
            if 'Trade Type' in factory_df.columns:
                factory_df = factory_df.drop(columns=['Trade Type'])
            if 'Shipper' in factory_df.columns:
                factory_df = factory_df.drop(columns=['Shipper'])
                
            factory_file_path = os.path.join(output_dir, f'reimport_invoice_factory_{factory}.xlsx')
            
            # Save factory invoice
            safe_save_to_excel(factory_df, factory_file_path)
    
    return result_df

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