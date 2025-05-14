import pandas as pd

# Define a function to check for Chinese characters
def contains_chinese(text):
    if pd.isna(text):
        return False
    for char in str(text):
        if '\u4e00' <= char <= '\u9fff':
            return True
    return False

# Read the reimport invoice
try:
    # First, check what sheets are available
    xls = pd.ExcelFile('outputs/reimport_invoice.xlsx')
    print(f"Available sheets: {xls.sheet_names}")

    # Try to read the second sheet with no header
    df = pd.read_excel('outputs/reimport_invoice.xlsx', sheet_name=1, header=None)

    # Print all rows to find where the actual data starts
    print("\nAll rows in the sheet (no header):")
    for i, row in df.iterrows():
        if i >= 8 and i <= 14:  # Focus on the data rows we saw earlier
            print(f"Row {i}: {row.tolist()}")

    # Now read the data with the correct header row (9)
    print("\nReading data with header row 9:")
    df_data = pd.read_excel('outputs/reimport_invoice.xlsx', sheet_name=1, header=9)

    # Print the column names
    print("\nColumns in the data:")
    print(df_data.columns.tolist())

    # Also check the RECI sheet
    print("\nChecking RECI sheet:")
    try:
        reci_df = pd.read_excel('outputs/reimport_invoice.xlsx', sheet_name=2, header=9)
        print("RECI sheet columns:")
        print(reci_df.columns.tolist())
        print("\nRECI sheet preview:")
        print(reci_df.head(6).to_string())

        # Check for Chinese characters in Commodity Description (Customs)
        if 'Commodity Description (Customs)' in reci_df.columns:
            chinese_count = reci_df['Commodity Description (Customs)'].apply(contains_chinese).sum()
            print(f"\nChinese Commodity Description count in RECI sheet: {chinese_count} out of {len(reci_df)} rows ({chinese_count/len(reci_df)*100:.1f}%)")

            if chinese_count > 0:
                print("\nRows with Chinese descriptions in RECI sheet:")
                chinese_rows = reci_df[reci_df['Commodity Description (Customs)'].apply(contains_chinese)]
                for _, row in chinese_rows.iterrows():
                    part_number = row.get('Part Number', 'N/A')
                    description = row.get('Commodity Description (Customs)', 'N/A')
                    print(f"Part Number: {part_number}, Description: {description}")
        else:
            print("Commodity Description (Customs) column not found in RECI sheet")
    except Exception as e:
        print(f"Error reading RECI sheet: {e}")

    # Print a preview of the data
    print("\nData preview:")
    print(df_data.head(6).to_string())

    # Check for Commodity Description column
    commodity_desc_col = 'Commodity Description (Customs)'
    print(f"\nUsing column: '{commodity_desc_col}'")

    # Check if the column exists
    if commodity_desc_col in df_data.columns:
        print(f"Column '{commodity_desc_col}' found in the data")
    else:
        # Try to find a similar column
        for col in df_data.columns:
            if 'Commodity' in str(col) or 'Description' in str(col):
                commodity_desc_col = col
                print(f"Found similar column: '{commodity_desc_col}'")
                break

    if commodity_desc_col:
        # Check for empty values
        empty_desc_count = df_data[commodity_desc_col].isna().sum()
        print(f"\nEmpty Commodity Description count: {empty_desc_count} out of {len(df_data)} rows ({empty_desc_count/len(df_data)*100:.1f}%)")

        # Check if Commodity Description contains Chinese characters
        def contains_chinese(text):
            if pd.isna(text):
                return False
            for char in str(text):
                if '\u4e00' <= char <= '\u9fff':
                    return True
            return False

        chinese_desc_count = df_data[commodity_desc_col].apply(contains_chinese).sum()
        print(f"Chinese Commodity Description count: {chinese_desc_count} out of {len(df_data)} rows ({chinese_desc_count/len(df_data)*100:.1f}%)")

        # Print rows with Chinese descriptions
        if chinese_desc_count > 0:
            print("\nRows with Chinese descriptions:")
            chinese_rows = df_data[df_data[commodity_desc_col].apply(contains_chinese)]
            for _, row in chinese_rows.iterrows():
                part_number = row.get('Part Number', 'N/A')
                description = row.get(commodity_desc_col, 'N/A')
                print(f"Part Number: {part_number}, Description: {description}")
    else:
        print("\nCommodity Description column not found in the data")

except Exception as e:
    print(f"Error reading reimport invoice: {e}")
