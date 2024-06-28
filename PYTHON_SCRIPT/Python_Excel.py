# iteration through excel file
import pandas as pd

# Load the Excel file
file_path = 'sample.xlsx'
sheet_name = 'Sheet1'

# Read the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Iterate through the table
for index, row in df.terrors():
    # Access data in each row
    print(f"Row {index}:")
    for col_name in df.columns:
        print(f"  {col_name}: {row[col_name]}")
