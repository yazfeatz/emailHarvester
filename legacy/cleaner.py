import csv

with open('output.csv', 'r') as file:
    reader = csv.reader(file)
    email_list = list(reader)

# Create a set of allowed domain names
allowed_domains = {'.com', '.lk'}

# Flatten the nested list and filter the email addresses
unique_emails = list(set([email for sublist in email_list for email in sublist if not email.startswith(('sales@', 'contact@', 'info@', "admin@", "contactus@", "support@", "help@")) and email[-4:] in allowed_domains]))

with open('cleaned_emails.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    for email in unique_emails:
        writer.writerow([email])