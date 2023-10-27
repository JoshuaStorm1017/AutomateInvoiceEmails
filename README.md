# AutomateInvoiceEmails
Automated Email System for Invoices (grabbing information from PDF file name)

AutomatedEmailScript This repository contains a Python script (AutomateEmails.py) for automating the process of sending emails.

The python code is in the master branch of this repository.

Description The script take emails from a list of customers from an Excel file (Customer_Emails_For_Invoices.xlsx), finds the corresponding PDF invoices in the 'Invoices_TO_BE_Emailed' folder, sends an email to each customer with the corresponding invoice attached, and finally, moves the sent invoices from the 'Invoices_TO_BE_Emailed' folder to the 'Invoices_THAT_WERE_EMailed' folder.

This was created to save me some time at work when we have a lot of invoices to be emailed as software does not do this for us. IMPORTANT This script retrieves the Invoice number and Customer number from the PDF file name. If your PDF files do not contain these, this scipt will not work.

Usage: You need to update the placeholders in the script with your information. Variables such as your_company_name, your_department, your_phone_number, your_email etc. should be replaced with your actual company name, department, phone number, company email. Also, you need to update parent_folder with the path where your invoice files and customer excel file are located.

Place your Customer_Emails_For_Invoices.xlsx file and the PDF invoices you want to email in the 'Invoices_TO_BE_Emailed' folder.

To run the Python script, follow the command:

python AutomatedEmails.py

Your Python environment needs to have all the necessary packages installed. Packages Needed: pandas pywin32

Feel Free to do whatever you please with this script.
