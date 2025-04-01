import pandas as pd

# Read the policy file
policy_df = pd.read_excel('testfiles/policy.xlsx')

# Print all information about the policy file
print("Policy File Information:")
print("\nColumns with their types:")
for col in policy_df.columns:
    print(f"Column name: '{col}', Type: {type(col)}, Repr: {repr(col)}")

print("\nFirst row values:")
for col in policy_df.columns:
    print(f"Column '{col}': {policy_df[col].iloc[0]}, Type: {type(policy_df[col].iloc[0])}")

# Print the entire DataFrame in a more readable format
print("\nFull DataFrame:")
print(policy_df.to_string())

# Try to save it as CSV for debugging
policy_df.to_csv('policy_debug.csv', encoding='utf-8')
print("\nSaved as CSV for easier inspection") 