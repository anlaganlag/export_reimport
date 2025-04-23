process_shipping_list.py break down

I'll help you understand this project by breaking it down into its main components and functionality.

This is a shipping list processing system that handles import/export invoices and packing lists. Let me explain the key aspects:

1. **Main Purpose**
The script processes shipping lists and generates two types of invoices:
- Export invoices (for general trade items)
- Re-import invoices (for items being reimported)
- Packing lists for both types

2. **Key Input Files**:
- A packing list Excel file (`original_packing_list.xlsx`)
- A policy file (`policy.xlsx`) containing configuration parameters like:
  - Exchange rates
  - Insurance rates
  - Markup percentages
  - Company information
  - Bank details

3. **Main Functions**:

a) `read_excel_file()`: 
- Handles multi-level headers (English + Chinese)
- Provides fallback mechanisms for different Excel formats

b) `process_shipping_list()`:
- The main processing function
- Converts source data into standardized formats
- Calculates prices and weights
- Generates multiple output files

c) `apply_excel_styling()`:
- Applies professional formatting to Excel outputs
- Sets column widths, fonts, borders
- Handles number formatting

4. **Key Calculations**:
- CIF Price (Cost, Insurance, and Freight)
- FOB Price (Free On Board)
- Insurance amounts
- Freight costs per kg
- Currency conversions (RMB to USD)

5. **Output Files**:
The script generates several Excel files in the `outputs` directory:
- `export_invoice.xlsx`: For general trade items
- `reimport_invoice.xlsx`: For re-import items
- `cif_original_invoice.xlsx`: Original CIF calculations
- Individual reimport invoices for different projects/factories

6. **Special Features**:
- Handles both Chinese and English column names
- Supports multiple trade types (general trade vs. buying trade)
- Merges cells in packing lists for better readability
- Converts numerical amounts to English words
- Splits data by project and factory

7. **Error Handling**:
- Robust error checking for file operations
- Fallback mechanisms for missing columns
- Detailed logging for debugging
- Retry mechanisms for file access issues

8. **Command Line Usage**:
```bash
python process_shipping_list.py --packing-list <path> --policy <path> --output-dir <path> [--debug]
```

Would you like me to dive deeper into any particular aspect of the project? For example:
1. The calculation logic
2. The file format requirements
3. The styling and formatting details
4. The error handling mechanisms
5. The project/factory splitting logic

Let me know what interests you most and I'll provide more detailed information!
