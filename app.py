# Insert your python code here# IMPORT PACKAGES IN THIS CELL

import os
import openpyxl
import pandas as pd

import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows# INSERT ANSWER FOR QUESTION 1 HERE
while True:    
    directory = single_input("Please input the directory where the ProductData.xlsx, MasterInvoice.pdf, and Order files are located.").strip('"')
    if directory.endswith('\\') == False:
        directory = directory + '\\'

    file_list = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx','Order4.xlsx'] 

    for file in file_list: 
        if os.path.exists(directory + file):
            continue
        else:
            show_message("Error", "The directory you entered does not contain the required files")
            break
    else:
        show_message('Thank you', 'All files are in the provided directory')
        break

print(directory)
# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

df = pd.read_excel(directory +'ProductData.xlsx', sheet_name = 'Descriptions')
for index, row in df.iterrows():
    description = row['Description']
    unit_number = row['Unit number']
    product_descriptions[unit_number] = description

df1 = pd.read_excel(directory +'ProductData.xlsx', sheet_name = 'Prices')
for index, row in df1.iterrows():
    ppu = row['Price per unit']
    unit_number = row['Unit number']
    product_prices[unit_number] = ppu

print(product_descriptions)
print(product_prices)
# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    template =openpyxl.load_workbook(directory + filename, data_only=True)
    sheet = template['OrderForm']

    order_date = sheet['C9'].value.strftime('%m/%d/%Y')
    email = str(sheet['C13'].value)
    name = str(sheet['C14'].value)
    company_name = str(sheet['C15'].value)
    address1 = str(sheet['C16'].value)
    city = str(sheet['C17'].value)
    state = str(sheet['C18'].value)
    zipcode = str(sheet['C19'].value)
    address2 =f'{city}, {state}, {zipcode}'
    product1 =sheet['B24'].value
    quantity1 =sheet['C24'].value
    product2 =sheet['B25'].value
    quantity2 =sheet['C25'].value
    product3 =sheet['B26'].value
    quantity3 =sheet['C26'].value

    products = [product1, product2, product3]
    quantities = [quantity1, quantity2, quantity3]

    order_data = {}
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = order_date
    order_data['CustomerAddress1'] = address1
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = company_name
    order_data['CustomerAddress2'] = address2

    total = 0
    i = 0

    for prod in products:
        qty = quantities[i]

        if pd.isnull(prod) == False:
            desc = product_descriptions[prod]
            price = product_prices[prod]
            amount = price * float(qty)

            order_data['Description' + str(i)] = str(desc)
            order_data['Price' + str(i)] = f"${price:,.2f}"
            order_data['Quantity' + str(i)] = str(qty)
            order_data['Amount' + str(i)] = f"${amount:.2f}"
            
            total += amount
        else:
            order_data['Description' + str(i)] = ""
            order_data['Price' + str(i)] = ""
            order_data['Quantity' + str(i)] = ""
            order_data['Amount' + str(i)] = ""
        i += 1    
    order_data['Total'] = f'${total :.2f}'

    return order_data, name, email, order_date 

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)


print("Order Data Dict:", order_data)
print("Name:", name)
print("Email:", email)
print("Date:", date)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    gmail_user = "jaren.laws@gmail.com"
    gmail_password = open('gmail_password.txt').read()
    recipients = [email]
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    body = f"""Dear {name},
    
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (810) 555-2372
    Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)
        

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'jaren.laws@gmail.com', '11/13/2023', directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num += 1
