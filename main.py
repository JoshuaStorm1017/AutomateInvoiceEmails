import os
import pandas as pd
import win32com.client as win32
import shutil
import logging

# Initialize logging
logging.basicConfig(filename='email_invoices.log', level=logging.INFO)

# Define company details
your_company_name = "Company ABC"
your_department = "Department XYZ"
your_phone_number = "123456789"
your_email = "you@yourcompany.com"

try:
    # Define the path to the parent folder
    parent_folder = r'C:\User\YOUR_USER_NAME\Automated_Emailed_Invoices'

    # Define the paths to the subfolders and Excel file
    to_be_emailed_folder = os.path.join(parent_folder, 'Invoices_TO_BE_Emailed')
    emailed_folder = os.path.join(parent_folder, 'Invoices_THAT_WERE_EMailed')
    excel_folder = os.path.join(parent_folder, 'Customer_Emails_For_Invoices')
    excel_file = 'Customer_Emails_For_Invoices.xlsx'

    # Define a dictionary specifying the desired data type for each column
    dtype_dict = {'Customer_Number': 'str'}

    # Read the Excel file into a pandas DataFrame with the specified data types
    df = pd.read_excel(os.path.join(excel_folder, excel_file), dtype=dtype_dict)

    # Get a list of all PDF files in the 'Invoices_TO_BE_Emailed' folder
    pdf_files = [f for f in os.listdir(to_be_emailed_folder) if f.endswith('.pdf')]

    # Initialize the Outlook application
    outlook = win32.Dispatch('outlook.application')

    # Loop through the PDF files
    for pdf_file in pdf_files:
        # Extract the customer number from the PDF file name
        customer_number, invoice_number, *_ = pdf_file.split('_')

        # Pad the customer number with leading zeros to make it six digits
        customer_number = customer_number.zfill(6)

        # Find the corresponding email address in the DataFrame
        email_address = df.loc[df['Customer_Number'] == customer_number, 'Customer_Email'].values

        # If the customer number exists in the DataFrame
        if len(email_address) > 0:
            # Create a new email
            mail = outlook.CreateItem(0)

            # Set the recipient, subject, and body of the email
            mail.To = email_address[0]
            mail.Subject = f'Invoice {invoice_number} from {your_company_name}'
            mail.Body = f'''Hello,

Attached is your invoice from {your_company_name}. Please do not hesitate to contact us with any questions or concerns.

Thank you and have a nice day,

{your_department}
{your_company_name}
P – {your_phone_number} 
{your_email}'''

            # Attach the PDF file to the email
            attachment = os.path.join(to_be_emailed_folder, pdf_file)
            mail.Attachments.Add(attachment)

            # Send the email
            mail.Send()

            # Log the successful email send
            logging.info(f"Successfully sent email for customer {customer_number}.")

            # Move the PDF file to the 'Invoices_THAT_WERE_EMailed' folder
            shutil.move(attachment, emailed_folder)
        else:
            logging.warning(f"No email address found for customer {customer_number}.")

    print('All invoices have been emailed.')

except Exception as e:
    logging.error(f"An error occurred: {e}")
