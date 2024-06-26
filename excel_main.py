import pandas as pd
from math import isnan
from pdf_main import create_pdf
from tkinter import Tk, messagebox  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
from work_data import address_map

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing


def select_file(prompt_message, file_name, file_type):
    file = askopenfilename(title=prompt_message, filetypes=[(file_name, file_type)])
    if file == '':
        quit()
    return file


# show an "Open" dialog box and return the path to the selected file
messagebox.showinfo(title='Excel To PDF', message='Please select an excel file to merge')
user_selected_file = select_file('Select Excel file', 'Excel File', '*.xlsx')

excel_file = pd.read_excel(user_selected_file)

messagebox.showinfo(title='Excel To PDF', message='Please select the PDF form.')
selected_pdf = select_file('Select PDF form', 'PDF form', '*.pdf')

columns_to_select = ['First Name:', 'Last Name:', 'Last 4:', 'Ext/Tel:', 'E-mail Address:', 'Years:', 'Building:',
                     'Employee Number', 'Address', 'City', 'State', 'ZipCode']
selected_columns = excel_file[columns_to_select]

list_of_objects = selected_columns.to_dict(orient='records')

count = 0

while count < len(list_of_objects):
    building_name = list_of_objects[count]['Building:']  # will return name of building
    list_of_objects[count]['Work_Address'] = address_map.get(building_name)  # looks up address via get
    list_of_objects[count].update(address_map['work_dict'])

    if isnan(list_of_objects[count]['Last 4:']):  # uses the math module
        list_of_objects[count]['Last 4:'] = ''
    else:
        list_of_objects[count]['Last 4:'] = int(list_of_objects[count]['Last 4:'])
        while len(str(list_of_objects[count]['Last 4:'])) != 4:
            formatted_ssn = str(list_of_objects[count]['Last 4:']).zfill(4)  # adds leading 0's
            list_of_objects[count]['Last 4:'] = formatted_ssn

    create_pdf(list_of_objects[count], selected_pdf)
    count += 1
