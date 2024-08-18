import pandas as pd

# Path to the Excel file
file_path = r'C:\Users\Lenovo\OneDrive\Desktop\Divyam\JD_Extractor\Unprocessed\RealEstate_Ahmedabad.xlsx'

# Read the Excel file
df = pd.read_excel(file_path)

# Select columns with indices 0, 3, and 6
selected_columns = df.iloc[:, [0, 13, 17]]

# Remove rows with any missing values in the selected columns
selected_columns = selected_columns.dropna()

# Rename the selected columns
new_column_names = ['Link', 'Name', 'Address']
selected_columns.columns = new_column_names

# Save the modified DataFrame to the same Excel file
# Use 'openpyxl' engine to write to Excel files
selected_columns.to_excel(file_path, index=False, engine='openpyxl')

print(f"Excel file '{file_path}' has been updated with the selected columns, removed empty rows, and new names.")
