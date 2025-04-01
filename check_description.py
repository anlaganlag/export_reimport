import pandas as pd
import numpy as np

# Read the original packing list
df = pd.read_excel('testfiles/original_packing_list.xlsx')

# Skip the header row which might be empty
df = df.iloc[1:]

# List possible columns that might contain descriptions
description_columns = []

# Find columns that might contain descriptions
for col in df.columns:
    col_str = str(col).lower()
    if any(term in col_str for term in ['品名', '货描', '描述', '物料名称', 'description']):
        description_columns.append(col)

print(f"Found {len(description_columns)} possible description columns:")
for i, col in enumerate(description_columns, 1):
    print(f"{i}. '{col}'")
    
    # Count non-empty values
    non_empty_count = df[col].notna().sum()
    print(f"   Non-empty values: {non_empty_count} out of {len(df)} ({non_empty_count/len(df)*100:.1f}%)")
    
    # Print first 5 non-empty values
    print("   First 5 non-empty values:")
    non_empty = df[df[col].notna()][col].head(5)
    for j, val in enumerate(non_empty, 1):
        print(f"     - {val}")
    print()

# Check if there's a column that might be better to use for DESCRIPTION
print("\nRecommended column for DESCRIPTION field:")
best_col = None
best_count = 0

for col in description_columns:
    count = df[col].notna().sum()
    if count > best_count:
        best_count = count
        best_col = col

if best_col:
    print(f"Best column: '{best_col}' with {best_count} non-empty values ({best_count/len(df)*100:.1f}%)")
else:
    print("No suitable column found for DESCRIPTION.") 