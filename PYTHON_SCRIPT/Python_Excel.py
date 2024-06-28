import pandas as pd
import os

# Get the path to the Downloads folder
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

# Define the file path
file_path = os.path.join(downloads_folder, 'Sample.xlsx')

# Step 1: Create the Excel file

# Sample data
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'City': ['New York', 'Los Angeles', 'Chicago']
}

# Create a DataFrame
df = pd.DataFrame(data)

# Write the DataFrame to an Excel file
df.to_excel(file_path, index=False, sheet_name='Sheet1')

# Step 2: Read and iterate through the Excel file

# Load the Excel file
sheet_name = 'Sheet1'

# Read the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Iterate through the table
for index, row in df.iterrows():
    # Access data in each row
    print(f"Row {index}:")
    for col_name in df.columns:
        print(f"  {col_name}: {row[col_name]}")
