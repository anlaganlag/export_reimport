#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Module for reading and parsing Excel input files.
"""

import pandas as pd
import re
import os
import numpy as np

def read_excel_file(file_path, skip=0):
    """
    Read an Excel file into a pandas DataFrame, handling various formats.
    This is extracted from the original process_shipping_list.py.

    Args:
        file_path (str): Path to the Excel file
        skip (int): Number of header rows to skip

    Returns:
        tuple: (DataFrame with the data, dict of metadata)
    """
    metadata = {'title': None, 'doc_number': None}
    
    try:
        # First, try to extract document title and number
        try:
            title_df = pd.read_excel(file_path, nrows=1, header=None)
            first_cell = title_df.iloc[0, 0] if not title_df.empty else None
            
            # Try to extract document number if it exists
            if first_cell and isinstance(first_cell, str):
                metadata['title'] = first_cell
                
                # Look for document number pattern (e.g., PL-20250418-0001)
                number_match = re.search(r'[A-Z]+-\d+-\d+', first_cell)
                if number_match:
                    metadata['doc_number'] = number_match.group(0)
                else:
                    # Try another pattern
                    number_match = re.search(r'\d{8}-\d+', first_cell)
                    if number_match:
                        metadata['doc_number'] = number_match.group(0)
        except Exception as e:
            print(f"Warning: Could not extract title/number: {e}")
        
        # Try reading with different settings until successful
        df = None
        errors = []
        
        # Attempt 1: With specified header and skiprows
        try:
            df = pd.read_excel(file_path, skiprows=skip)
            # Verify we got actual data
            if df.shape[0] <= 1 or df.shape[1] <= 2:
                errors.append("Attempt 1 resulted in too few rows/columns")
                df = None
        except Exception as e:
            errors.append(f"Attempt 1 failed: {e}")
            
        # Attempt 2: Try with header=None if previous failed
        if df is None:
            try:
                df = pd.read_excel(file_path, header=None)
                # Find the actual header row
                for i in range(min(10, df.shape[0])):
                    potential_header = df.iloc[i]
                    # Check if this row looks like a header
                    if any(str(val).lower() in ['序号', 'no', 'part', '零件号', 'description', '描述'] 
                          for val in potential_header if val is not None):
                        # Use this row as header
                        df.columns = potential_header
                        df = df.iloc[i+1:].reset_index(drop=True)
                        break
            except Exception as e:
                errors.append(f"Attempt 2 failed: {e}")
        
        # Attempt 3: Try reading with Excel's built-in header detection
        if df is None:
            try:
                df = pd.read_excel(file_path)
            except Exception as e:
                errors.append(f"Attempt 3 failed: {e}")
                raise ValueError(f"Could not read Excel file {file_path}. Errors: {', '.join(errors)}")
        
        # Clean up column names
        if not df.empty:
            # Remove unnamed columns and convert column names to strings
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df.columns = [str(col).strip() if col is not None else f"Column_{i}" 
                         for i, col in enumerate(df.columns)]
            
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            # Replace NaN with None in string columns
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].apply(lambda x: None if pd.isna(x) else x)
                    
            # Attempt to identify and remove summary rows
            # This will be handled separately by model/transformer.py
            
        return df, metadata
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        raise

def detect_header_row(df):
    """
    Detect which row is likely to be the header row in a DataFrame.
    
    Args:
        df (DataFrame): Input DataFrame with potential header rows
        
    Returns:
        int: Index of the likely header row, or 0 if none found
    """
    # Common header fields to look for
    header_keywords = ['序号', 'no', 'part', '零件', '描述', 'description', 
                     '数量', 'quantity', '单位', 'unit']
    
    for i in range(min(10, len(df))):
        row = df.iloc[i]
        row_values = [str(v).lower() if v is not None else '' for v in row]
        
        # Count how many keywords are in this row
        keyword_matches = sum(1 for kw in header_keywords 
                            if any(kw in str(v).lower() for v in row_values))
        
        # If we have multiple matches, this is likely a header row
        if keyword_matches >= 3:
            return i
            
    return 0 