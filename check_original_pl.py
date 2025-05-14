import pandas as pd

# Read the original packing list
try:
    # First, check what sheets are available
    xls = pd.ExcelFile('testfiles/original_packing_list.xlsx')
    print(f"Available sheets: {xls.sheet_names}")
    
    # Try to read the first sheet, skipping the first 2 rows as done in the main script
    df = pd.read_excel('testfiles/original_packing_list.xlsx', skiprows=2)
    
    # Print the column names
    print("Columns in the original packing list:")
    print(df.columns.tolist())
    
    # Check if '进口清关货描' column exists
    if '进口清关货描' in df.columns:
        print("\n'进口清关货描' column found in the original packing list")
        
        # Print the values in the '进口清关货描' column
        print("\nValues in the '进口清关货描' column:")
        for i, value in enumerate(df['进口清关货描'].values):
            print(f"Row {i+1}: {value}")
            
        # Check if the values are in English
        def contains_chinese(text):
            if pd.isna(text):
                return False
            for char in str(text):
                if '\u4e00' <= char <= '\u9fff':
                    return True
            return False
        
        chinese_count = df['进口清关货描'].apply(contains_chinese).sum()
        print(f"\nChinese characters in '进口清关货描' column: {chinese_count} out of {len(df)} rows ({chinese_count/len(df)*100:.1f}%)")
        
        # Print the mapping between Part Number and 进口清关货描
        print("\nMapping between Part Number and 进口清关货描:")
        mapping = {}
        for _, row in df.iterrows():
            if '料号' in df.columns and '进口清关货描' in df.columns:
                part_number = row['料号']
                customs_desc = row['进口清关货描']
                mapping[part_number] = customs_desc
                print(f"Part Number: {part_number}, 进口清关货描: {customs_desc}")
    else:
        print("\n'进口清关货描' column not found in the original packing list")
        
except Exception as e:
    print(f"Error reading original packing list: {e}")
