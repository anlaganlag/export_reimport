#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Shim script to maintain backward compatibility with existing code.
This script will eventually replace process_shipping_list.py once refactoring is complete.
"""

from shipping_processor import process_shipping_list

# Re-export the process_shipping_list function to maintain API compatibility
__all__ = ['process_shipping_list']

if __name__ == "__main__":
    import sys
    import os
    
    # Handle command-line arguments similar to the original script
    if len(sys.argv) < 3:
        print("Usage: python process_shipping_list.py <packing_list_file> <policy_file> [output_dir]")
        sys.exit(1)
    
    # Parse command line arguments
    packing_list_file = sys.argv[1]
    policy_file = sys.argv[2]
    output_dir = sys.argv[3] if len(sys.argv) > 3 else "outputs"
    
    # Call the implementation
    process_shipping_list(packing_list_file, policy_file, output_dir) 