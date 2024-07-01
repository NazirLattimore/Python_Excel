import pandas as pd
import os

# Get the path to the Downloads folder
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

# Define the file path
file_path = os.path.join(downloads_folder, 'Sample.xlsx')

# Check if the file exists
if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
else:
    # Load the Excel file
    df = pd.read_excel(file_path, sheet_name='Sheet1')  # Adjust sheet_name as necessary

    # Access columns directly and iterate through index
    for idx in df.index:
        # Access data using column names
        name = df.at[idx, 'Name']
        age = df.at[idx, 'Age']
        city = df.at[idx, 'City']
        
        # Print row information
        print(f"Row {idx}:")
        print(f"  Name: {name}, Age: {age}, City: {city}")
