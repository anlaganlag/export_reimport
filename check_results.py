import pandas as pd

# Read the export invoice
df = pd.read_excel('outputs/export_invoice.xlsx')

# Print a preview of important columns
print("EXPORT INVOICE PREVIEW:")
preview_cols = ['NO.', 'Material code', 'DESCRIPTION', 'Qty', 'FOB单价', 'FOB总价']
print(df[preview_cols].head(10).to_string())

# Check for empty values in DESCRIPTION
empty_desc_count = df['DESCRIPTION'].isna().sum()
print(f"\nEmpty DESCRIPTION count: {empty_desc_count} out of {len(df)} rows ({empty_desc_count/len(df)*100:.1f}%)")

# Show some statistics
print("\nDescription Statistics:")
desc_lengths = df['DESCRIPTION'].str.len().dropna()
if not desc_lengths.empty:
    print(f"  Average length: {desc_lengths.mean():.1f} characters")
    print(f"  Min length: {desc_lengths.min()} characters")
    print(f"  Max length: {desc_lengths.max()} characters")

# Print a summary of non-empty columns
print("\nColumn Fill Rates:")
for col in df.columns:
    non_empty_count = df[col].notna().sum()
    fill_percentage = non_empty_count / len(df) * 100
    print(f"  {col}: {fill_percentage:.1f}% filled ({non_empty_count}/{len(df)})")

# Print some examples of descriptions
print("\nSample Descriptions:")
sample_desc = df[df['DESCRIPTION'].notna()]['DESCRIPTION'].sample(min(5, df['DESCRIPTION'].notna().sum()))
for i, desc in enumerate(sample_desc, 1):
    print(f"  {i}. {desc}") 