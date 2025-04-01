import pandas as pd

# Check the original packing list structure
print("Original Packing List Structure:")
packing_list_df = pd.read_excel('testfiles/original_packing_list.xlsx')
print("Columns:", list(packing_list_df.columns))
print("\nFirst few rows:")
print(packing_list_df.head())

# Check the policy file structure
print("\n\nPolicy File Structure:")
policy_df = pd.read_excel('testfiles/policy.xlsx')
print("Columns:", list(policy_df.columns))
# Print exact column names for debugging
print("\nExact Column Names and Their Types:")
for col in policy_df.columns:
    print(f"Column name: '{col}', Type: {type(col)}, Repr: {repr(col)}")
print("\nAll rows:")
print(policy_df) 