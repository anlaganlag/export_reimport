# -*- coding: utf-8 -*-
import openpyxl
import copy
import os
import sys

def merge_three_excel_files(first_file, middle_file, last_file, output_file):
    """
    Merge exactly three Excel files with specific requirements:
    - First file (h.xlsx) is used entirely
    - Middle file - second sheet is used for merging, first sheet is kept intact
    - Last file (f.xlsx) is used entirely
    - All merged cells and formatting are preserved
    """
    print(f"Merging files in order: {first_file}, {middle_file} (2nd sheet only), {last_file}")
    print(f"Output will be saved to: {output_file}")
    
    # Create a new workbook for the merged result
    merged_wb = openpyxl.Workbook()
    
    # Create Commercial Invoice sheet for merged content
    merged_sheet = merged_wb.active
    merged_sheet.title = 'Commercial Invoice'
    
    # Load the middle workbook to copy its first sheet
    try:
        middle_wb = openpyxl.load_workbook(middle_file, data_only=True)
        if len(middle_wb.sheetnames) > 0:
            # Create a new sheet in merged workbook for the first sheet from middle file
            packing_list_sheet = merged_wb.create_sheet('Packing List')
            
            # Copy the first sheet from middle workbook
            source_sheet = middle_wb[middle_wb.sheetnames[0]]
            
            # Copy all data and formatting from first sheet
            for row_idx, row in enumerate(source_sheet.rows, 1):
                for col_idx, source_cell in enumerate(row, 1):
                    target_cell = packing_list_sheet.cell(row=row_idx, column=col_idx)
                    
                    # Handle merged cells
                    if isinstance(source_cell, openpyxl.cell.cell.MergedCell):
                        # Skip setting value as it will be handled by the main cell
                        pass
                    else:
                        target_cell.value = source_cell.value
                        
                        # Copy styling if available
                        if hasattr(source_cell, '_style'):
                            target_cell._style = copy.copy(source_cell._style)
                            target_cell.font = copy.copy(source_cell.font)
                            target_cell.border = copy.copy(source_cell.border)
                            target_cell.fill = copy.copy(source_cell.fill)
                            target_cell.number_format = copy.copy(source_cell.number_format)
                            target_cell.protection = copy.copy(source_cell.protection)
                            target_cell.alignment = copy.copy(source_cell.alignment)
            
            # Copy merged cells from first sheet
            for merged_range in source_sheet.merged_cells:
                merge_range = f"{openpyxl.utils.get_column_letter(merged_range.min_col)}{merged_range.min_row}:{openpyxl.utils.get_column_letter(merged_range.max_col)}{merged_range.max_row}"
                try:
                    packing_list_sheet.merge_cells(merge_range)
                except ValueError as e:
                    print(f"Warning: Could not merge cells {merge_range} in Packing List: {e}")
                    
            # Copy column dimensions
            for col_idx, column in enumerate(source_sheet.columns, 1):
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                if col_letter in source_sheet.column_dimensions:
                    packing_list_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
                    
            # Copy row dimensions
            for row_idx in range(1, source_sheet.max_row + 1):
                if row_idx in source_sheet.row_dimensions:
                    packing_list_sheet.row_dimensions[row_idx].height = source_sheet.row_dimensions[row_idx].height
                    
            print(f"Successfully copied first sheet from {middle_file}")
    except Exception as e:
        print(f"Warning: Could not copy first sheet from {middle_file}: {e}")
    
    # Store workbooks and relevant sheets for merging
    file_data = []
    
    # First file (h.xlsx)
    try:
        wb1 = openpyxl.load_workbook(first_file, data_only=True)
        file_data.append((wb1, wb1.active, first_file))
    except Exception as e:
        print(f"Error opening first file {first_file}: {e}")
        return False
    
    # Middle file (only second sheet)
    try:
        wb2 = openpyxl.load_workbook(middle_file, data_only=True)
        # Get the second sheet (index 1)
        if len(wb2.sheetnames) > 1:
            sheet2 = wb2[wb2.sheetnames[1]]  # Second sheet
            file_data.append((wb2, sheet2, middle_file))
        else:
            print(f"Warning: Middle file {middle_file} does not have a second sheet")
            return False
    except Exception as e:
        print(f"Error opening middle file {middle_file}: {e}")
        return False
    
    # Last file (f.xlsx)
    try:
        wb3 = openpyxl.load_workbook(last_file, data_only=True)
        file_data.append((wb3, wb3.active, last_file))
    except Exception as e:
        print(f"Error opening last file {last_file}: {e}")
        return False
    
    # Merge all sheets' data with formatting
    row_offset = 0
    
    for wb, sheet, filename in file_data:
        print(f"Processing file for merge: {filename}")
        
        # Copy data and formatting
        for row_idx, row in enumerate(sheet.rows, 1):
            for col_idx, source_cell in enumerate(row, 1):
                target_cell = merged_sheet.cell(row=row_offset + row_idx, column=col_idx)
                
                # Handle merged cells
                if isinstance(source_cell, openpyxl.cell.cell.MergedCell):
                    # Skip setting value as it will be handled by the main cell
                    pass
                else:
                    target_cell.value = source_cell.value
                    
                    # Copy styling if available
                    if hasattr(source_cell, '_style'):
                        target_cell._style = copy.copy(source_cell._style)
                        target_cell.font = copy.copy(source_cell.font)
                        target_cell.border = copy.copy(source_cell.border)
                        target_cell.fill = copy.copy(source_cell.fill)
                        target_cell.number_format = copy.copy(source_cell.number_format)
                        target_cell.protection = copy.copy(source_cell.protection)
                        target_cell.alignment = copy.copy(source_cell.alignment)
        
        # Process merged cells
        for merged_range in sheet.merged_cells:
            min_row = merged_range.min_row + row_offset
            max_row = merged_range.max_row + row_offset
            min_col = merged_range.min_col
            max_col = merged_range.max_col
            
            # Create the new merge range
            merge_range = f"{openpyxl.utils.get_column_letter(min_col)}{min_row}:{openpyxl.utils.get_column_letter(max_col)}{max_row}"
            
            try:
                merged_sheet.merge_cells(merge_range)
                
                # Copy the value from the top-left cell of the source merged range
                source_value = sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
                merged_sheet.cell(row=min_row, column=min_col).value = source_value
            except ValueError as e:
                print(f"Warning: Could not merge cells {merge_range}: {e}")
        
        # Update row offset for the next file
        row_offset += sheet.max_row
    
    # Make sure the Packing List sheet is first in the workbook
    # Move sheets to correct order: Packing List first, then Commercial Invoice
    if 'Packing List' in merged_wb.sheetnames:
        merged_wb._sheets = [merged_wb['Packing List'], merged_wb['Commercial Invoice']]

    # Save the merged workbook
    try:
        merged_wb.save(output_file)
        print(f"Successfully saved merged file to: {output_file}")
        return True
    except Exception as e:
        print(f"Error saving output file {output_file}: {e}")
        return False

if __name__ == "__main__":
    # Check command line arguments
    if len(sys.argv) < 5:
        print("Usage: python merge.py <first_file.xlsx> <middle_file.xlsx> <last_file.xlsx> <output_file.xlsx>")
        sys.exit(1)
    
    # Get absolute paths for all files
    first_file = os.path.abspath(sys.argv[1])
    middle_file = os.path.abspath(sys.argv[2])
    last_file = os.path.abspath(sys.argv[3])
    output_file = os.path.abspath(sys.argv[4])
    
    # Perform the merge
    success = merge_three_excel_files(first_file, middle_file, last_file, output_file)
    
    if not success:
        print("Merge operation failed!")
        sys.exit(1)

# 在Windows系统下自动打开合并后的Excel文件
# if os.name == 'nt':
#     os.startfile(output_file)
#     print("Opening merged Excel file...")
#     print("按回车键退出程序...")
#     input()