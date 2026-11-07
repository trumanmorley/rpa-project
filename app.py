# Insert your python code here
# INSERT ANSWER FOR QUESTION 1 HERE
neccessary_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:
    directory = single_input("Please Input Directory File Path")

    if directory[-1] != '\\':
        directory = directory + '\\'

    correct_file_count = 0
    filenames = os.listdir(directory)
    for file in filenames:
        if file in neccessary_files:
            correct_file_count += 1

    if correct_file_count == 6:
        break
    else:
        show_message('Error', 'Incorrect File Path, Please Try Again!')

# INSERT ANSWER FOR QUESTION 2 HERE
df1 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Descriptions')

product_descriptions = {}
for i, rows in df1.iterrows():
    unit_number = rows['Unit number']
    description = rows['Description']
    product_descriptions[unit_number] = description

df2 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Prices')

product_prices ={}
for i, rows in df2.iterrows():
    unit_number = rows['Unit number']
    prices = rows['Price per unit']
    product_prices[unit_number] = prices


# INSERT ANSWER FOR QUESTION 3 HERE
import datetime as dt


def get_order_data(invoice_num, filename, directory):
    order_data = openpyxl.load_workbook(directory + filename)
    sheets = order_data['OrderForm']
    date = sheets['C9'].value
    email = sheets['C13'].value
    name = sheets['C14'].value
    total = 0

    new_dict = {}
    new_dict['InvoiceNum'] = str(invoice_num)
    new_dict['CustomerAddress1'] = sheets['C16'].value
    new_dict['CustomerAdress2'] = f'{sheets['C17'].value}, {sheets['C18'].value}, {sheets['C19'].value}'
    new_dict['Date'] = date
    new_dict['CustomerName'] = name
    new_dict['CustomerCompany'] = sheets['C15'].value

    if not sheets['B24'].value is None:
        first_price = product_prices[sheets['B24'].value]
        first_quantity = sheets['C24'].value
        first_amount = first_price * first_quantity
        new_dict['1Price'] = f'${first_price:,.2f}'
        new_dict['1Quantity'] = str(first_quantity)
        new_dict['1Amount'] = f'${first_amount:,.2f}'
        new_dict['Description'] = product_descriptions[sheets['B24'].value]
        total += first_amount

    if not sheets['B25'].value is None:
        sec_price = product_prices[sheets['B25'].value]
        sec_quantity = sheets['C25'].value
        sec_amount = sec_price * sec_quantity
        new_dict['2Price'] = f'${sec_price:,.2f}'
        new_dict['2Quantity'] = str(sec_quantity)
        new_dict['2Amount'] = f'${sec_amount:,.2f}'
        new_dict['Description'] = product_descriptions[sheets['B25'].value]
        total += first_amount
    
    if not sheets['B26'].value is None:
        third_price = product_prices[sheets['B26'].value]
        third_quantity = sheets['C26'].value
        third_amount = third_price * third_quantity
        new_dict['3Price'] = f'${third_price:,.2f}'
        new_dict['3Quantity'] = str(third_quantity)
        new_dict['3Amount'] = f'${third_amount:,.2f}'
        new_dict['Description'] = product_descriptions[sheets['B25'].value]
        total += first_amount

    if not sheets['B27'].value is None:
        four_price = product_prices[sheets['B27'].value]
        four_quantity = sheets['C27'].value
        four_amount = four_price * four_quantity
        new_dict['4Price'] = f'${four_price:,.2f}'
        new_dict['4Quantity'] = str(four_quantity)
        new_dict['4Amount'] = f'${four_amount:,.2f}'
        new_dict['Description'] = product_descriptions[sheets['B27'].value]
        total += first_amount

    if not sheets['B28'].value is None:
        five_price = product_prices[sheets['B28'].value]
        five_quantity = sheets['C28'].value
        five_amount = five_price * five_quantity
        new_dict['5Price'] = f'${five_price:,.2f}'
        new_dict['5Quantity'] = str(five_quantity)
        new_dict['5Amount'] = f'${five_amount:,.2f}'
        new_dict['Description'] = product_descriptions[sheets['B28'].value]
        total += first_amount

    new_dict['Total'] = f'${total:,.2f}'
    return new_dict, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(order_data)
print(name)
print(email)
print(date)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    pdf_template = directory + 'MasterInvoice.pdf'
    new_pdf = directory + f'Invoice {invoice_num}.pdf'
    create_pdf(pdf_template, new_pdf, order_data)

test = create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    subject = f'Your Order from PythonTech Co - Invoice {invoice_num}'
    gmail_user = 'trumanmorley@gmail.com'
    passcode = open('C:/Users/Test/OneDrive - Brigham Young University/ACC525/17 - Project 3 - Robotic Process Automation/gmail_password.txt').read()
    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84602
Phone: (801) 555-2372
Email: sales@pythontechco.com
    """

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = 'trumanmorley@gmail.com'
    message['Subject'] = subject
    message.set_content(body)

    message = add_attachment(directory + f'Invoice {invoice_num}.pdf', message)

    server = smtplib.SMTP_SSL('smtp.gmail.com')
    server.login(gmail_user, passcode)
    server.send_message(message)
    server.quit

send_email(8515, name, 'trumanmorley@gmail.com', date, directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)

for files in filenames:
    order = 8515
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(order, file, directory)
        create_invoice(order, order_data, directory)
        send_email(order, name, email, date)
        order += 1