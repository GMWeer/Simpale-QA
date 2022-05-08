#!/usr/bin/env python
# coding: utf-8

# In[24]:


import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')
#sg.theme('DarkBlue9')

EXCEL_FILE = (r'C:\Users\HP i5\Downloads\mm.xlsx')
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Date', size=(15,1)), sg.InputText(key='Date')],
    [sg.Text('Sub_name', size=(15,1)), sg.InputText(key='Sub_name')],
    [sg.Text('Fdr_name', size=(15,1)), sg.Combo(['Kandy', 'Colombo', 'Galle', 'Kaluthara', 'Badulla'],key='Fdr_name')],
    [sg.Text('Time off', size=(15,1)), sg.InputText(key='Time off')],
    [sg.Text('Time on', size=(15,1)), sg.InputText(key='Time on')],
    #[sg.Text('Favorite Colour', size=(15,1)), sg.Combo(['Green', 'Blue', 'Red', 'Orange', '1'], key='Favorite Colour')],
    [sg.Text('Delay hrs.', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='Delay hrs.')],
    [sg.Text('Line_section', size=(15,1)),
                            sg.Checkbox('Section 1', key='Section 1'),
                            sg.Checkbox('Section 2', key='Section 2'),
                            sg.Checkbox('Section 1 & 2', key='Section 1 & 2')],
   
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()


# In[ ]:





# In[ ]:





# In[ ]:




