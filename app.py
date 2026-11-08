# Insert your python code here

# INSERT ANSWER FOR QUESTION 1 HERE
def get_directory():
    directory = select_file("Please select the file that is in the directory that you want to use.")
    split_directory = directory.split("/")
    del split_directory[-1]
    directory = "/".join(split_directory)
    directory = directory + "/"
    return (directory)


while True:   
    directory = get_directory() 
    filenames = os.listdir(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and'Order4.xlsx' in filenames :
        break
    else:
        show_message("Error", "The Directory you entered is invalid. Please try again.")

print(directory)


# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

file_data = directory + "ProductData.xlsx"

def exceltodict (file, sheet, colname):
    df = pd.read_excel(file, sheet) 
    dictionary = {}
    for index,row in df.iterrows():
        dictionary[row["Unit number"]] = row[colname]

    return (dictionary)


product_descriptions = exceltodict(file_data, "Descriptions", "Description" )
product_prices = exceltodict(file_data, "Prices", "Price per unit")

display(product_descriptions)
product_prices



# INSERT ANSWER FOR QUESTION 3 HERE
# import datetime as dt
from analyticskit import *
df = get_pdf('MasterInvoice.pdf')
df



def get_order_data (invoice_num, filename, directory):
    path = directory + "ProductData.xlsx"

    template = openpyxl.load_workbook(filename, data_only=True)
    selected_sheet = template['OrderForm']

    custName = selected_sheet['C14'].value
    custEmail = selected_sheet['C13'].value
    dtOrder = selected_sheet['C9'].value
    dtOrder = str(dtOrder)
    dtOrder = dtOrder.split()
    dtOrder = dtOrder[0]
    dtOrder = dtOrder.split("-")
    dtOrder = f"{dtOrder[1]}/{dtOrder[2]}/{dtOrder[0]}"
    
    update_dict = {}
    update_dict['InvoiceNum'] = str(invoice_num)
    update_dict['Date'] = dtOrder
    update_dict['CustomerAddress1'] = selected_sheet['C16'].value
    update_dict['CustomerName'] = selected_sheet['C14'].value
    update_dict['CustomerCompany'] = selected_sheet['C15'].value
    City = selected_sheet['C17'].value
    State = selected_sheet['C18'].value
    Zipcode = selected_sheet['C19'].value
    update_dict['CustomerAddress2'] = f"{City}, {State} {Zipcode}"
    
    def countitems ():
        
        try: 
            if selected_sheet["C25"].value > 0:
                try: 
                    selected_sheet["C26"].value > 0
                    num_items = 3
                except:
                    num_items = 2
        except:
            num_items = 1
        return (num_items)


    def fixnum (num):
        numstr = f'${num:,.2f}'
        return (numstr)
    

    thecount = countitems()
    
    if thecount == 1:
        Price0 = product_prices[selected_sheet['B24'].value]
        Quantity0 = selected_sheet['C24'].value
        Amount0 = Quantity0 * Price0
        update_dict['Price0'] = fixnum(Price0)
        update_dict['Quantity0'] = Quantity0
        update_dict['Amount0'] = fixnum(Amount0)
        update_dict['Description0'] = product_descriptions[selected_sheet['B24'].value]
        
        update_dict['Total'] = fixnum(Amount0)
    elif thecount == 2:
        Price0 = product_prices[selected_sheet['B24'].value]
        Quantity0 = selected_sheet['C24'].value
        Amount0 = Quantity0 * Price0
        update_dict['Price0'] = fixnum(Price0)
        update_dict['Quantity0'] = Quantity0
        update_dict['Amount0'] = fixnum(Amount0)
        update_dict['Description0'] = product_descriptions[selected_sheet['B24'].value]

        Price1 = product_prices[selected_sheet['B25'].value]
        Quantity1 = selected_sheet['C25'].value
        Amount1 = Quantity1 * Price1
        update_dict['Price1'] = fixnum(Price1)
        update_dict['Quantity1'] = Quantity1
        update_dict['Amount1'] = fixnum(Amount1)
        update_dict['Description1'] = product_descriptions[selected_sheet['B25'].value]
        
        update_dict['Total'] = fixnum(Amount0 + Amount1)
    elif thecount == 3:
        Price0 = product_prices[selected_sheet['B24'].value]
        Quantity0 = selected_sheet['C24'].value
        Amount0 = Quantity0 * Price0
        update_dict['Price0'] = fixnum(Price0)
        update_dict['Quantity0'] = Quantity0
        update_dict['Amount0'] = fixnum(Amount0)
        update_dict['Description0'] = product_descriptions[selected_sheet['B24'].value]

        Price1 = product_prices[selected_sheet['B25'].value]
        Quantity1 = selected_sheet['C25'].value
        Amount1 = Quantity1 * Price1
        update_dict['Price1'] = fixnum(Price1)
        update_dict['Quantity1'] = Quantity1
        update_dict['Amount1'] = fixnum(Amount1)
        update_dict['Description1'] = product_descriptions[selected_sheet['B25'].value]

        Price2 = product_prices[selected_sheet['B26'].value]
        Quantity2 = selected_sheet['C26'].value
        Amount2 = Quantity2 * Price2
        update_dict['Price2'] = fixnum(Price2)
        update_dict['Quantity2'] = Quantity2
        update_dict['Amount2'] = fixnum(Amount2)
        update_dict['Description2'] = product_descriptions[selected_sheet['B26'].value]
        
        update_dict['Total'] = fixnum(Amount0 + Amount1 + Amount2)
    
    return (update_dict, custName, custEmail, dtOrder)


order_data, customerName, customerEmail, dateOrder = get_order_data(invoice_num=8515, filename="Order1.xlsx", directory=directory)


display(customerName)
display(customerEmail)
display(dateOrder)
display(order_data)



# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice (invoice_num,order_data,directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)
    
create_invoice(invoice_num=8515,order_data=order_data,directory=directory)



# INSERT ANSWER FOR QUESTION 5 HERE
gmail_password_path = select_file('Please select the gmail password file')
def send_email (invoice_num,name,email,date,directory):
    # Enter your gmail user name and password

    gmail_user = email
    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = ["coltonbissegger@gmail.com"]

    # Email subect

    subject = f"Your Order from PythonTech Co. - Invoice {invoice_num}"

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    
    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(invoice_num=8515,name=customerName,email=customerEmail,date=dateOrder,directory=directory)




# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames.sort()
filenames

invnum = 8515
for fn in filenames:
    if "Order" in fn and "xlsx" in fn:
        order_data, customerName, customerEmail, dateOrder = get_order_data(invoice_num=invnum, filename=fn, directory=directory)
        create_invoice(invoice_num=invnum,order_data=order_data,directory=directory)
        send_email(invoice_num=invnum,name=customerName,email=customerEmail,date=dateOrder,directory=directory)
        invnum = invnum + 1
