import os
import re
from pandas import read_excel, DataFrame

# Define regular expression to match email addresses
email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[a-zA-Z]{2,}\b'

# Set the path to the folder containing the Excel files
folder_path = '../sources'

# Create an empty list to store all email addresses
all_emails = []

# Loop through all files in the folder
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)

    # Check if the file is an Excel file
    if filename.lower().endswith(('.xlsx', '.xls')):
        try:
            # Read the Excel file into a pandas dataframe
            df = read_excel(file_path)

            # Loop through all columns in the dataframe
            for col in df.columns:
                # Loop through all values in the column
                for value in df[col]:
                    # Check if the value is a string
                    if isinstance(value, str):
                        # Use regex to find all email addresses in the value
                        emails = re.findall(email_regex, value)

                        # Append the email addresses found to the list
                        all_emails.extend(emails)

        except Exception as e:
            print(f"Error reading file '{file_path}': {e}")

# Convert the list of email addresses to a pandas dataframe
email_df = DataFrame({'Email': all_emails})

# Write the dataframe to a CSV file
output_path = '../output.csv'
try:
    mode = 'a' if os.path.exists(output_path) else 'w'
    email_df.to_csv(output_path, mode=mode, index=False, header=not os.path.exists(output_path))
    print(f'Successfully extracted {len(all_emails)} email addresses to {output_path}.')
except Exception as e:
    print(f"Error writing to file '{output_path}': {e}")
