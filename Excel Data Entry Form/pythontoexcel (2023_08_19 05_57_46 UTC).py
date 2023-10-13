#Excel to Python using PySimpleGUI and openpyxl

import PySimpleGUI as sg
import pandas as pd

sg.theme('Black')

EXCEL_FILE = 'data1.xlsx'
df = pd.read_excel(EXCEL_FILE)

# each list within the Layout List represents a column in the GUI
# pySimpleGUI has different "Elements". On the first row below you have a simple header text.
# In the 2nd row there is another text with the size where it's 15 characters wide, and 1 character tall (height).

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('Favorite Color', size=(15,1)), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Color')],
    [sg.Text('Marital Status', size=(15,1)),
        sg.Checkbox('Single', key='Single'),
        sg.Checkbox('Married', key='Married'),
        sg.Checkbox('Divorced', key='Divorced'),
        sg.Checkbox('Widowed', key='Widowed')],
    [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                      initial_value=0, key='Children')],
    [sg.Submit(), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
window.close()