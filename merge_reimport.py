#!/usr/bin/env python3
import sys
import os
from openpyxl import load_workbook, Workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter
from copy import copy

def copy_cell_format(source_cell, target_cell):
    """Copy formatting from source cell to target cell."""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def copy_column_widths(source_sheet, target_sheet):
    """Copy column widths from source sheet to target sheet."""
    for column in source_sheet.column_dimensions:
        target_sheet.column_dimensions[column] = copy(source_sheet.column_dimensions[column])

def copy_merged_cells(source_sheet, target_sheet, row_offset=0):
    """Copy merged cells from source sheet to target sheet with row offset."""
    # Get all merged ranges from source sheet
    merged_ranges = []
    for merged_range in source_sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
        
        # Apply row offset
        min_row += row_offset
        max_row += row_offset
        
        # Create the new merged range
        new_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        merged_ranges.append((new_range, min_col, min_row))
    
    # Get the value from the top-left cell of each merged range
    for merged_range, min_col, min_row in merged_ranges:
        try:
            # Get the value from the top-left cell
            source_value = source_sheet.cell(row=min_row-row_offset, column=min_col).value
            
            # Merge the cells in the target sheet
            target_sheet.merge_cells(merged_range)
            
            # Set the value in the top-left cell
            target_sheet.cell(row=min_row, column=min_col).value = source_value
        except Exception as e:
            print(f"Warning: Could not merge cells {merged_range}: {e}")

def copy_sheet_content(source_sheet, target_sheet, row_offset=0):
    """Copy content from source sheet to target sheet with row offset."""
    # Get all merged ranges from source sheet
    merged_cells = set()
    for merged_range in source_sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                merged_cells.add((row, col))
    
    # Copy cell values and formatting
    for row in source_sheet.rows:
        for cell in row:
            # Skip merged cells (they will be handled separately)
            if (cell.row, cell.column) not in merged_cells:
                target_cell = target_sheet.cell(row=cell.row + row_offset, column=cell.column)
                target_cell.value = cell.value
                copy_cell_format(cell, target_cell)

def merge_sheets(header_file, main_file, footer_file, output_file, sheet_name, output_wb=None):
    """Merge header, main content, and footer into a single sheet."""
    try:
        # Load workbooks
        header_wb = load_workbook(header_file)
        main_wb = load_workbook(main_file)
        footer_wb = load_workbook(footer_file)
        
        # Get the sheets
        header_sheet = header_wb.active
        main_sheet = main_wb[sheet_name]
        footer_sheet = footer_wb.active
        
        # Create or get output workbook and sheet
        if output_wb is None:
            output_wb = Workbook()
            output_sheet = output_wb.active
            output_sheet.title = sheet_name
        else:
            # Create a new sheet in the existing workbook
            output_sheet = output_wb.create_sheet(sheet_name)
        
        # Get row counts
        header_max_row = header_sheet.max_row
        main_max_row = main_sheet.max_row
        
        # Copy header content and merged cells
        copy_sheet_content(header_sheet, output_sheet)
        copy_merged_cells(header_sheet, output_sheet)
        
        # Copy main content and merged cells
        copy_sheet_content(main_sheet, output_sheet, header_max_row)
        copy_merged_cells(main_sheet, output_sheet, header_max_row)
        
        # Copy footer content and merged cells
        copy_sheet_content(footer_sheet, output_sheet, header_max_row + main_max_row)
        copy_merged_cells(footer_sheet, output_sheet, header_max_row + main_max_row)
        
        # Copy column widths from all sheets
        copy_column_widths(header_sheet, output_sheet)
        copy_column_widths(main_sheet, output_sheet)
        copy_column_widths(footer_sheet, output_sheet)
        
        return output_wb
    except Exception as e:
        print(f"Error merging sheets: {e}")
        return None

def main():
    if len(sys.argv) != 7:
        print("Usage: python merge_reimport.py <input_file> <output_file> <pl_header_file> <pl_footer_file> <ci_header_file> <ci_footer_file>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    pl_header_file = sys.argv[3]
    pl_footer_file = sys.argv[4]
    ci_header_file = sys.argv[5]
    ci_footer_file = sys.argv[6]
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file {input_file} does not exist")
        sys.exit(1)
    
    # Load the input workbook to get sheet names
    try:
        input_wb = load_workbook(input_file)
        sheet_names = input_wb.sheetnames
    except Exception as e:
        print(f"Error loading input file {input_file}: {e}")
        sys.exit(1)
    
    # Create output workbook
    output_wb = None
    
    # Process each sheet
    for sheet_name in sheet_names:
        print(f"Processing sheet: {sheet_name}")
        
        # Determine which header and footer files to use
        if sheet_name == 'Packing List':
            header_file = pl_header_file
            footer_file = pl_footer_file
        else:
            header_file = ci_header_file
            footer_file = ci_footer_file
        
        # Skip if header or footer files are missing or empty
        if not header_file or not footer_file:
            print(f"Warning: Header or footer files not provided for sheet {sheet_name}, skipping")
            continue
            
        if not os.path.exists(header_file) or not os.path.exists(footer_file):
            print(f"Warning: Header or footer files missing for sheet {sheet_name}, skipping")
            continue
        
        # Merge the sheets
        try:
            output_wb = merge_sheets(header_file, input_file, footer_file, output_file, sheet_name, output_wb)
            if output_wb:
                print(f"Successfully merged sheet: {sheet_name}")
            else:
                print(f"Failed to merge sheet: {sheet_name}")
                sys.exit(1)
        except Exception as e:
            print(f"Error merging sheet {sheet_name}: {e}")
            sys.exit(1)
    
    # Save the output workbook
    try:
        output_wb.save(output_file)
        print(f"Successfully processed all sheets in {input_file}")
    except Exception as e:
        print(f"Error saving output file: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()