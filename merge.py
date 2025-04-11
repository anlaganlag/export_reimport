# -*- coding: utf-8 -*-
import openpyxl
import copy
import os
import sys

def copy_cell_formatting(source_cell, target_cell):
    """Helper function to copy cell formatting."""
    if hasattr(source_cell, '_style'):
        target_cell._style = copy.copy(source_cell._style)
        target_cell.font = copy.copy(source_cell.font)
        target_cell.border = copy.copy(source_cell.border)
        target_cell.fill = copy.copy(source_cell.fill)
        target_cell.number_format = copy.copy(source_cell.number_format)
        target_cell.protection = copy.copy(source_cell.protection)
        target_cell.alignment = copy.copy(source_cell.alignment)

def copy_sheet(source_sheet, target_sheet):
    """Copy contents and formatting from source sheet to target sheet."""
    # Copy cell values and formatting
    for row_idx, row in enumerate(source_sheet.rows, 1):
        for col_idx, source_cell in enumerate(row, 1):
            target_cell = target_sheet.cell(row=row_idx, column=col_idx)
            
            if not isinstance(source_cell, openpyxl.cell.cell.MergedCell):
                target_cell.value = source_cell.value
                copy_cell_formatting(source_cell, target_cell)
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells:
        merge_range = f"{openpyxl.utils.get_column_letter(merged_range.min_col)}{merged_range.min_row}:{openpyxl.utils.get_column_letter(merged_range.max_col)}{merged_range.max_row}"
        try:
            target_sheet.merge_cells(merge_range)
        except ValueError as e:
            print(f"Warning: Could not merge cells {merge_range}: {e}")
    
    # Copy column dimensions
    for col_idx, column in enumerate(source_sheet.columns, 1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        if col_letter in source_sheet.column_dimensions:
            target_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
    
    # Copy row dimensions
    for row_idx in range(1, source_sheet.max_row + 1):
        if row_idx in source_sheet.row_dimensions:
            target_sheet.row_dimensions[row_idx].height = source_sheet.row_dimensions[row_idx].height

def append_sheet_with_offset(source_sheet, target_sheet, row_offset, filename):
    """Append source sheet to target sheet with row offset."""
    # Copy data with offset
    for row_idx, row in enumerate(source_sheet.rows, 1):
        for col_idx, source_cell in enumerate(row, 1):
            target_cell = target_sheet.cell(row=row_offset + row_idx, column=col_idx)
            
            if not isinstance(source_cell, openpyxl.cell.cell.MergedCell):
                target_cell.value = source_cell.value
                copy_cell_formatting(source_cell, target_cell)
    
    # Process merged cells with offset
    for merged_range in source_sheet.merged_cells:
        min_row = merged_range.min_row + row_offset
        max_row = merged_range.max_row + row_offset
        min_col = merged_range.min_col
        max_col = merged_range.max_col
        
        merge_range = f"{openpyxl.utils.get_column_letter(min_col)}{min_row}:{openpyxl.utils.get_column_letter(max_col)}{max_row}"
        
        try:
            target_sheet.merge_cells(merge_range)
            source_value = source_sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
            target_sheet.cell(row=min_row, column=min_col).value = source_value
        except ValueError as e:
            print(f"Warning: Could not merge cells {merge_range}: {e}")
    
    return source_sheet.max_row

def apply_column_widths(sheet, column_widths, default_widths=None):
    """Apply column widths to sheet based on header names."""
    default_widths = default_widths or {
        'Material code': 35,
        'Unit Price': 20,
        'DESCRIPTION': 30,
        'Model NO.': 20,
        'default': 15
    }
    
    for col_idx, cell in enumerate(sheet[1], 1):
        col_name = cell.value
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        
        if col_name in column_widths:
            sheet.column_dimensions[col_letter].width = column_widths[col_name]
        elif col_name in default_widths:
            sheet.column_dimensions[col_letter].width = default_widths[col_name]
        else:
            sheet.column_dimensions[col_letter].width = default_widths['default']

def load_workbook_safely(file_path):
    """Load workbook with error handling."""
    try:
        return openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:
        print(f"Error opening file {file_path}: {e}")
        return None

def merge_three_excel_files(first_file, middle_file, last_file, output_file):
    """
    Merge exactly three Excel files with specific requirements:
    - First file used entirely
    - Middle file - second sheet for merging, first sheet preserved
    - Last file used entirely
    - All merged cells and formatting preserved
    """
    print(f"Merging files: {first_file}, {middle_file}, {last_file}")
    
    # Create output workbook
    merged_wb = openpyxl.Workbook()
    merged_sheet = merged_wb.active
    merged_sheet.title = 'Commercial Invoice'
    
    # Process middle file first to extract both sheets
    middle_wb = load_workbook_safely(middle_file)
    if not middle_wb or len(middle_wb.sheetnames) < 2:
        print(f"Error: Middle file must have at least 2 sheets")
        return False
        
    # Create and copy Packing List from middle file's first sheet
    packing_list_sheet = merged_wb.create_sheet('Packing List')
    copy_sheet(middle_wb[middle_wb.sheetnames[0]], packing_list_sheet)
    print(f"Copied first sheet from {middle_file}")
    
    # Extract column widths from middle file's second sheet
    middle_ci_sheet = middle_wb[middle_wb.sheetnames[1]]
    column_widths = {}
    for col_idx, cell in enumerate(middle_ci_sheet[1], 1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        if cell.value and col_letter in middle_ci_sheet.column_dimensions:
            column_widths[cell.value] = middle_ci_sheet.column_dimensions[col_letter].width
    
    # Prepare data for merging Commercial Invoice
    files_to_merge = [
        (first_file, lambda wb: wb.active),
        (middle_file, lambda wb: wb[wb.sheetnames[1]]),
        (last_file, lambda wb: wb.active)
    ]
    
    # Merge sheets vertically
    row_offset = 0
    for file_path, sheet_selector in files_to_merge:
        wb = load_workbook_safely(file_path)
        if not wb:
            return False
            
        sheet = sheet_selector(wb)
        print(f"Processing: {file_path}")
        row_offset += append_sheet_with_offset(sheet, merged_sheet, row_offset, file_path)
    
    # Apply column widths
    apply_column_widths(merged_sheet, column_widths)
    
    # Reorder sheets
    merged_wb._sheets = [merged_wb['Packing List'], merged_wb['Commercial Invoice']]
    
    # Save result
    try:
        merged_wb.save(output_file)
        print(f"Successfully saved merged file to: {output_file}")
        return True
    except Exception as e:
        print(f"Error saving output file {output_file}: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("Usage: python merge.py <first_file.xlsx> <middle_file.xlsx> <last_file.xlsx> <output_file.xlsx>")
        sys.exit(1)
    
    files = [os.path.abspath(sys.argv[i]) for i in range(1, 5)]
    success = merge_three_excel_files(*files)
    
    if not success:
        print("Merge operation failed!")
        sys.exit(1)

# 在Windows系统下自动打开合并后的Excel文件
# if os.name == 'nt':
#     os.startfile(output_file)
#     print("Opening merged Excel file...")
#     print("按回车键退出程序...")
#     input()