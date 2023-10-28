# Automated Invoice Email System

## Overview
This project was developed to optimize the invoicing system within my current employment by automating the process of emailing invoices. The manual task of sending out numerous invoices was time-consuming and inefficient. This automation initiative has streamlined the invoicing process, making it more efficient and freeing up time for other tasks.

## Features

- Reads customer emails from an Excel file (`Customer_Emails_For_Invoices.xlsx`).
- Finds corresponding PDF invoices in the 'Invoices_TO_BE_Emailed' folder.
- Sends an email to each customer with the corresponding invoice attached.
- Moves sent invoices to the 'Invoices_THAT_WERE_Emailed' folder for record-keeping.
- Retrieves Invoice number and Customer number from the PDF file names.

## How to Use

1. Clone the repository from [GitHub](https://github.com/JoshuaStorm1017/AutomatedEmailScript/tree/master).
2. Place your `Customer_Emails_For_Invoices.xlsx` file and PDF invoices in the 'Invoices_TO_BE_Emailed' folder.
3. Update placeholders in the script (`main.py`) with your specific company information.
4. Run the following command to execute the script:
   python main.py


## Prerequisites

- Python environment with the following packages installed:
- pandas
- pywin32

## Technologies

- Python

