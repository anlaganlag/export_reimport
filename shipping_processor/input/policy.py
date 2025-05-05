#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Module for reading and parsing policy files.
"""

import pandas as pd
import re
import numpy as np

def read_policy_file(policy_file):
    """
    Read policy file to extract configuration parameters.
    This is extracted from the original process_shipping_list.py.

    Args:
        policy_file (str): Path to the policy Excel file

    Returns:
        dict: Dictionary of policy parameters
    """
    try:
        # Initialize policy data dictionary
        policy_data = {
            'exchange_rate': None,
            'markup_rate': None,
            'insurance_rate': None,
            'doc_number': None,
            'company_name': None,
            'company_address': None,
            'bank_name': None,
            'account_no': None,
            'swift_code': None,
            'branch_address': None,
            'factory_info': {}
        }
        
        # Read policy file
        df = pd.read_excel(policy_file)
        
        # Extract policy number from title or header
        # Try first row
        if not df.empty and isinstance(df.iloc[0, 0], str):
            first_cell = df.iloc[0, 0]
            # Look for policy number pattern
            number_match = re.search(r'Policy\s+No[.:]\s*([A-Z0-9-]+)', first_cell, re.IGNORECASE)
            if number_match:
                policy_data['doc_number'] = number_match.group(1)
            else:
                # Try another pattern (e.g. PL-20250418-0001)
                number_match = re.search(r'[A-Z]+-\d+-\d+', first_cell)
                if number_match:
                    policy_data['doc_number'] = number_match.group(0)
                else:
                    # Try yet another pattern
                    number_match = re.search(r'\d{8}-\d+', first_cell)
                    if number_match:
                        policy_data['doc_number'] = number_match.group(0)
        
        # If not found in title, try to find in columns
        if not policy_data['doc_number']:
            # Look for column with '编号' or 'number'
            number_cols = [col for col in df.columns if '编号' in str(col).lower() or 'number' in str(col).lower()]
            if number_cols and not df[number_cols[0]].empty:
                # Get the first non-NA value
                number_values = df[number_cols[0]].dropna()
                if not number_values.empty:
                    policy_data['doc_number'] = str(number_values.iloc[0])
        
        # Extract rates from the dataframe
        # Exchange rate
        exchange_rates = extract_rate(df, ['汇率', 'exchange rate', 'exchange'])
        if exchange_rates:
            policy_data['exchange_rate'] = exchange_rates[0]
            
        # Markup rate
        markup_rates = extract_rate(df, ['加价率', 'markup rate', 'markup'])
        if markup_rates:
            policy_data['markup_rate'] = markup_rates[0]
            
        # Insurance rate
        insurance_rates = extract_rate(df, ['保险费率', 'insurance rate', 'insurance'])
        if insurance_rates:
            policy_data['insurance_rate'] = insurance_rates[0]
        
        # Extract company and bank information
        company_info = extract_company_info(df)
        if company_info:
            policy_data.update(company_info)
            
        # Extract factory information if available
        factory_info = extract_factory_info(df)
        if factory_info:
            policy_data['factory_info'] = factory_info
            
        return policy_data
    except Exception as e:
        print(f"Error reading policy file {policy_file}: {e}")
        raise

def extract_rate(df, keywords):
    """
    Extract rate values from a dataframe based on keywords.
    
    Args:
        df (DataFrame): The policy dataframe
        keywords (list): List of keywords to look for
        
    Returns:
        list: List of extracted rate values
    """
    # First, try to find exact column matches
    matching_cols = []
    for col in df.columns:
        if col is not None and any(keyword.lower() in str(col).lower() for keyword in keywords):
            matching_cols.append(col)
    
    if matching_cols:
        # Extract values from matching columns
        rates = []
        for col in matching_cols:
            # Get non-NaN, non-zero values
            values = df[col].dropna()
            numeric_values = pd.to_numeric(values, errors='coerce')
            valid_values = numeric_values[numeric_values != 0].dropna()
            if not valid_values.empty:
                rates.extend(valid_values.tolist())
        
        return rates
    
    # If no columns found, try row-based search
    for _, row in df.iterrows():
        for col, val in row.items():
            if val is not None and any(keyword.lower() in str(val).lower() for keyword in keywords):
                # Look for numeric values in this row
                for col2, val2 in row.items():
                    if col2 != col and val2 is not None:
                        try:
                            rate = float(val2)
                            if rate > 0:
                                return [rate]
                        except (ValueError, TypeError):
                            continue
    
    return []

def extract_company_info(df):
    """
    Extract company and bank information from policy file.
    
    Args:
        df (DataFrame): The policy dataframe
        
    Returns:
        dict: Dictionary with company and bank info
    """
    info = {}
    
    # Keywords to look for
    company_keywords = ['company name', '公司名称', 'company']
    address_keywords = ['company address', '公司地址', 'address']
    bank_keywords = ['bank name', '银行名称', 'bank']
    account_keywords = ['account', '账号', 'account no']
    swift_keywords = ['swift', 'swift code', 'swift编码']
    branch_keywords = ['branch', 'branch address', '支行地址']
    
    # Function to find value based on keywords
    def find_value_by_keywords(row, keywords):
        for col, val in row.items():
            if val is not None and any(keyword.lower() in str(val).lower() for keyword in keywords):
                # Return the next cell value if possible
                col_idx = df.columns.get_loc(col)
                if col_idx + 1 < len(df.columns):
                    next_val = row.iloc[col_idx + 1]
                    if pd.notna(next_val) and str(next_val).strip():
                        return next_val
        return None
    
    # Iterate through rows to find information
    for _, row in df.iterrows():
        # Company name
        if 'company_name' not in info or not info['company_name']:
            company_name = find_value_by_keywords(row, company_keywords)
            if company_name:
                info['company_name'] = str(company_name).strip()
        
        # Company address
        if 'company_address' not in info or not info['company_address']:
            company_address = find_value_by_keywords(row, address_keywords)
            if company_address:
                info['company_address'] = str(company_address).strip()
        
        # Bank name
        if 'bank_name' not in info or not info['bank_name']:
            bank_name = find_value_by_keywords(row, bank_keywords)
            if bank_name:
                info['bank_name'] = str(bank_name).strip()
        
        # Account number
        if 'account_no' not in info or not info['account_no']:
            account_no = find_value_by_keywords(row, account_keywords)
            if account_no:
                info['account_no'] = str(account_no).strip()
        
        # Swift code
        if 'swift_code' not in info or not info['swift_code']:
            swift_code = find_value_by_keywords(row, swift_keywords)
            if swift_code:
                info['swift_code'] = str(swift_code).strip()
        
        # Branch address
        if 'branch_address' not in info or not info['branch_address']:
            branch_address = find_value_by_keywords(row, branch_keywords)
            if branch_address:
                info['branch_address'] = str(branch_address).strip()
    
    return info

def extract_factory_info(df):
    """
    Extract factory-specific information from policy file.
    
    Args:
        df (DataFrame): The policy dataframe
        
    Returns:
        dict: Dictionary with factory information
    """
    factory_info = {}
    
    # Look for factory section
    factory_section = False
    factory_name = None
    
    for _, row in df.iterrows():
        row_values = [str(val).lower() if val is not None else '' for val in row]
        row_text = ' '.join(row_values)
        
        # Check if we've hit the factory section
        if 'factory' in row_text or '工厂' in row_text:
            factory_section = True
            continue
        
        if not factory_section:
            continue
            
        # Look for factory name and project pattern
        for i, val in enumerate(row):
            if val is not None and isinstance(val, str):
                # Check for factory name
                if 'factory' in val.lower() or '工厂' in val.lower():
                    # Get factory name from next cell if possible
                    if i + 1 < len(row) and pd.notna(row.iloc[i + 1]):
                        factory_name = str(row.iloc[i + 1]).strip()
                    elif ':' in val or '：' in val:
                        # Try to extract name after colon
                        factory_name = re.split(r'[:：]', val)[-1].strip()
                    
                    if factory_name and factory_name not in factory_info:
                        factory_info[factory_name] = {'projects': []}
                        
                # Check for project pattern
                project_match = re.search(r'project[:\s]+([A-Za-z0-9_-]+)', str(val), re.IGNORECASE)
                if project_match and factory_name:
                    project = project_match.group(1).strip()
                    if project not in factory_info[factory_name]['projects']:
                        factory_info[factory_name]['projects'].append(project)
    
    return factory_info 