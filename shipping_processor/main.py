#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Main entry point for the shipping processor.
This module maintains the original API while we gradually refactor.
"""

import os
import sys
import importlib.util
import pandas as pd

# First attempt to import from the original module if it exists
original_module_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'process_shipping_list.py')
if os.path.exists(original_module_path):
    # Dynamically import the original module
    spec = importlib.util.spec_from_file_location("original_process_shipping_list", original_module_path)
    original_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(original_module)
    
    # Use the original implementation for now
    def process_shipping_list(packing_list_file, policy_file, output_dir='outputs'):
        """
        Main function to process a shipping list.
        This is a thin wrapper around the original implementation for now.
        
        Args:
            packing_list_file (str): Path to the packing list Excel file
            policy_file (str): Path to the policy Excel file 
            output_dir (str): Directory to save output files
            
        Returns:
            None
        """
        try:
            # Attempt to use refactored input modules as a test
            from shipping_processor.input import read_excel_file, read_policy_file
            from shipping_processor.model import translate_unit, merge_india_invoice_rows
            from shipping_processor.output import safe_save_to_excel
            from shipping_processor.format import apply_excel_styling
            
            # Log that we're using refactored modules
            print("INFO: Testing refactored modules along with original implementation")
            
            # But still call the original implementation
            return original_module.process_shipping_list(packing_list_file, policy_file, output_dir)
        except ImportError as e:
            print(f"WARNING: Could not import refactored modules: {e}")
            return original_module.process_shipping_list(packing_list_file, policy_file, output_dir)
        except Exception as e:
            print(f"ERROR in refactored modules: {e}")
            # Fall back to original implementation
            return original_module.process_shipping_list(packing_list_file, policy_file, output_dir)
else:
    # If original module doesn't exist, implement the function here
    # This will be our final implementation once refactoring is complete
    def process_shipping_list(packing_list_file, policy_file, output_dir='outputs'):
        """
        Main function to process a shipping list.
        
        Args:
            packing_list_file (str): Path to the packing list Excel file
            policy_file (str): Path to the policy Excel file 
            output_dir (str): Directory to save output files
            
        Returns:
            None
        """
        # Import necessary modules
        from shipping_processor.input import read_excel_file, read_policy_file
        from shipping_processor.model import translate_unit, merge_india_invoice_rows, split_by_project_and_factory
        from shipping_processor.output import safe_save_to_excel, merge_packing_list_cells
        from shipping_processor.format import apply_excel_styling, apply_pl_footer_styling
        
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"Processing shipping list: {packing_list_file}")
        print(f"Using policy file: {policy_file}")
        print(f"Output directory: {output_dir}")
        
        # Step 1: Read input files
        packing_list_df, pl_metadata = read_excel_file(packing_list_file)
        policy_data = read_policy_file(policy_file)
        
        # TODO: Continue implementation with the refactored modules
        
        print("Refactored implementation not yet complete")
        print("Using modules:")
        print(f"- Input: read_excel_file, read_policy_file")
        print(f"- Model: translate_unit, merge_india_invoice_rows, split_by_project_and_factory")
        print(f"- Output: safe_save_to_excel, merge_packing_list_cells")
        print(f"- Format: apply_excel_styling, apply_pl_footer_styling")
        
        # For now, raise not implemented, but in the future this will be complete
        raise NotImplementedError("Refactored implementation not yet complete") 