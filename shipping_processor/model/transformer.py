#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Module for transforming DataFrame data.
"""

import pandas as pd
import numpy as np
import re

def merge_india_invoice_rows(df):
    """
    Merge rows in an invoice for India factory.
    This is extracted from the original process_shipping_list.py.

    Args:
        df (DataFrame): Input DataFrame
        
    Returns:
        DataFrame: Transformed DataFrame with merged rows
    """
    # Check if DataFrame is empty
    if df.empty:
        return df
        
    # Create a copy to avoid modifying the original
    result_df = df.copy()
    
    # Check for serial number column
    sn_col = None
    part_col = None
    desc_col = None
    qty_col = None
    
    # Identify important columns
    for col in df.columns:
        col_lower = str(col).lower()
        if 'no' in col_lower or 'sn' in col_lower or '序号' in col_lower:
            sn_col = col
        elif 'part' in col_lower or '零件' in col_lower:
            part_col = col
        elif 'desc' in col_lower or '描述' in col_lower:
            desc_col = col
        elif 'qty' in col_lower or 'quantity' in col_lower or '数量' in col_lower:
            qty_col = col
    
    # If any required column is missing, return the original DataFrame
    if not all([sn_col, part_col, desc_col, qty_col]):
        return df
    
    # Group rows by part number and aggregate
    grouped = result_df.groupby(part_col, as_index=False).agg({
        # Keep the first serial number
        sn_col: 'first',
        # Concatenate descriptions with newlines or commas
        desc_col: lambda x: '\n'.join(str(val) for val in x if val) if len(x) > 1 else x.iloc[0],
        # Sum quantities
        qty_col: 'sum'
    })
    
    # For other columns, use appropriate aggregation
    for col in result_df.columns:
        if col in [sn_col, part_col, desc_col, qty_col]:
            continue
            
        col_lower = str(col).lower()
        if any(weight in col_lower for weight in ['weight', 'gross', 'net', '重量']):
            # Sum weights
            grouped[col] = result_df.groupby(part_col)[col].sum().values
        elif any(price in col_lower for price in ['price', 'unit price', '单价']):
            # Keep first price
            grouped[col] = result_df.groupby(part_col)[col].first().values
        elif any(amount in col_lower for amount in ['amount', 'total', '金额']):
            # Sum amounts
            grouped[col] = result_df.groupby(part_col)[col].sum().values
        else:
            # For other columns, keep the first value
            grouped[col] = result_df.groupby(part_col)[col].first().values
    
    # Renumber serial numbers
    for i, idx in enumerate(grouped.index):
        grouped.at[idx, sn_col] = i + 1
        
    return grouped

def split_by_project_and_factory(df):
    """
    Split a DataFrame by project and factory.
    This is extracted from the original process_shipping_list.py.

    Args:
        df (DataFrame): Input DataFrame
        
    Returns:
        dict: Dictionary of {factory/project: DataFrame}
    """
    # Check if DataFrame is empty
    if df.empty:
        return {'default': df}
    
    # Create a copy to avoid modifying the original
    result_df = df.copy()
    
    # Try to identify factory or project column
    factory_col = None
    project_col = None
    
    # Identify important columns based on common names
    for col in df.columns:
        col_lower = str(col).lower()
        if 'factory' in col_lower or '工厂' in col_lower:
            factory_col = col
        elif 'project' in col_lower or '项目' in col_lower:
            project_col = col
    
    # If no factory or project column, use part number prefix
    if not factory_col and not project_col:
        # Try to identify part number column
        part_col = None
        for col in df.columns:
            col_lower = str(col).lower()
            if 'part' in col_lower or '零件' in col_lower:
                part_col = col
                break
        
        if part_col:
            # Extract project/factory code from part number
            def extract_prefix(part_num):
                if not part_num or not isinstance(part_num, str):
                    return 'unknown'
                # Try to extract prefix (usually 2-4 letters at the beginning)
                match = re.match(r'^([A-Za-z]{2,4})', str(part_num))
                if match:
                    return match.group(1).upper()
                return 'unknown'
            
            # Add a new column for factory/project based on part number prefix
            result_df['factory_project'] = result_df[part_col].apply(extract_prefix)
            factory_col = 'factory_project'
    
    # Split by factory or project
    if factory_col:
        # Group by factory
        grouped = dict(tuple(result_df.groupby(factory_col)))
        
        # Clean up groups
        result = {}
        for key, group_df in grouped.items():
            if key is None or pd.isna(key):
                result['unknown'] = group_df.drop(factory_col, axis=1) if factory_col == 'factory_project' else group_df
            else:
                result[str(key)] = group_df.drop(factory_col, axis=1) if factory_col == 'factory_project' else group_df
                
        return result
    elif project_col:
        # Group by project
        grouped = dict(tuple(result_df.groupby(project_col)))
        
        # Clean up groups
        result = {}
        for key, group_df in grouped.items():
            if key is None or pd.isna(key):
                result['unknown'] = group_df
            else:
                result[str(key)] = group_df
                
        return result
    else:
        # No split possible
        return {'all': result_df}

def find_column_with_pattern(df, patterns, target_col_name=None):
    """
    Find columns matching a list of patterns.
    This is extracted from the original process_shipping_list.py.

    Args:
        df (DataFrame): Input DataFrame
        patterns (list): List of patterns to match
        target_col_name (str): Optional name for the new column
        
    Returns:
        tuple: (column name if found, index of matching pattern)
    """
    if df.empty:
        return None, -1
        
    # Convert patterns to lowercase for case-insensitive matching
    patterns = [p.lower() if isinstance(p, str) else str(p).lower() for p in patterns]
    
    # Look for exact column name matches first
    for col in df.columns:
        col_lower = str(col).lower()
        for i, pattern in enumerate(patterns):
            if pattern == col_lower:
                return col, i
    
    # Then look for partial matches
    for col in df.columns:
        col_lower = str(col).lower()
        for i, pattern in enumerate(patterns):
            if pattern in col_lower:
                return col, i
    
    # If target column name is specified, create it
    if target_col_name and target_col_name not in df.columns:
        # Return None to indicate no existing column was found
        return None, -1
        
    return None, -1 