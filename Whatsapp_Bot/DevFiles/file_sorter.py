import pandas as pd

# Load the Excel file
df = pd.read_excel(r'C:\Users\Lenovo\OneDrive\Desktop\Divyam\Whatsapp_Bot_final\Akshat_1set.xlsx')

# 1. Remove rows where 'Number' column does not start with '+91'
df = df[df['Number'].str.startswith('+91')]

# 2. Remove '+91' and any spaces from the 'Number' column
df['Number'] = df['Number'].str.replace('+91', '').str.replace(' ', '', regex=True)

# 3. Remove '~' from the 'Name' column
df['Name'] = df['Name'].str.replace('~', '', regex=True)

# 4. Remove duplicate rows based on the 'Number' column
df = df.drop_duplicates(subset='Number')

import pdb;pdb.set_trace()


# Save the cleaned DataFrame to a new Excel file (optional)
df.to_excel('cleaned_file.xlsx', index=False)

print("Data cleaning complete!")
