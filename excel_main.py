import pandas as pd
from math import isnan
from pdf_main import create_pdf
from tkinter import Tk, messagebox  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
messagebox.showinfo(title='Excel To PDF', message='Please select an excel file to merge')

user_selected_file = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
if user_selected_file == '':
    quit()

messagebox.showinfo(title='Excel To PDF', message='Please select the PDF form.')
selected_pdf = askopenfilename()
if selected_pdf == '':
    quit()


def add_leading_zeroes(value):
    while len(value) != 4:
        value = '0' + value
    return value


work_forch_address = '1300 Morris Park Avenue'
work_ken_address = '1410 Pelham Parkway South'
work_price_address = '1301 Morris Park Avenue'
work_van_address = '1225 Morris Park Avenue'
work_dict = {'Work_City': 'Bronx',
             'Work_State': 'NY',
             'Work_Zipcode': '10461'}

excel_file = pd.read_excel(user_selected_file)

columns_to_select = ['First Name:', 'Last Name:', 'Last 4:', 'Ext/Tel:', 'E-mail Address:', 'Years:', 'Building:',
                     'Employee Number', 'Address', 'City', 'State', 'ZipCode']
selected_columns = excel_file[columns_to_select]

list_of_objects = selected_columns.to_dict(orient='records')

count = 0

while count < len(list_of_objects):
    if list_of_objects[count]['Building:'] == 'Kennedy':
        list_of_objects[count]['Work_Address'] = work_ken_address
    elif list_of_objects[count]['Building:'] == 'Price':
        list_of_objects[count]['Work_Address'] = work_price_address
    elif list_of_objects[count]['Building:'] == 'Van Etten':
        list_of_objects[count]['Work_Address'] = work_van_address
    else:
        list_of_objects[count]['Work_Address'] = work_forch_address

    list_of_objects[count].update(work_dict)

    if isnan(list_of_objects[count]['Last 4:']):  # uses the math module
        list_of_objects[count]['Last 4:'] = ''
    else:
        list_of_objects[count]['Last 4:'] = int(list_of_objects[count]['Last 4:'])
        while len(str(list_of_objects[count]['Last 4:'])) != 4:
            formatted_ssn = add_leading_zeroes(str(list_of_objects[count]['Last 4:']))
            list_of_objects[count]['Last 4:'] = formatted_ssn

    create_pdf(list_of_objects[count], selected_pdf)
    count += 1
