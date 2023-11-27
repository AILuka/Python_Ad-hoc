#!/usr/bin/env python
# coding: utf-8

# In[1]:


print('Loading...')


# In[2]:


import pandas as pd
from tkinter import *
import tkinter as ttk
import os


# ### Inputs window

# In[3]:


#get data from entry field, workbook name
def show_message():
    global directory
    global workbook_name 
    directory = entry.get()
    workbook_name = entry2.get()
    root.destroy()

# create a window with input of workbookname
root = Tk()
root.title("Inputs")
root.geometry("400x150+500+300") 

lbl_1 = Label(root, text='Enter a file directory:')
lbl_1.place(x = 10, y = 10) 

entry = ttk.Entry(width=60)
entry.place(x = 10, y = 30) 
#defauld entry text
entry.insert(0, os.getcwd())

lbl_2 = Label(root, text='Enter an original file name:')
lbl_2.place(x = 10, y = 60) 

entry2 = ttk.Entry(width=60)
entry2.place(x = 10, y = 80)
#defauld entry text
entry2.insert(0, 'in_staff_raw_data.xlsx')

btn = ttk.Button(text="Start", command=show_message)
btn.place(x = 10, y = 110) 

root.mainloop()


# In[4]:


print('Doing the math, wait a minute...')


# ### Calculation

# In[5]:


#create a python directiry format
directory = directory.replace("\\","/")+'/'
file_inputs = 'Inputs.xlsx'
#upload files
inputs_fx = pd.read_excel(directory+file_inputs, sheet_name='FX', header=0)
inputs_ranges = pd.read_excel(directory+file_inputs, sheet_name='Ranges', header=0)
staff_raw_data = pd.read_excel(directory+workbook_name, header=1)


# In[6]:


#create a table of FX incl. all currencies for correction with LE
list_fx_le = inputs_fx[inputs_fx['Legal Entity'].notnull()]
#create a list incl. all currencies for correction without LE
list_fx_wo_le = inputs_fx[inputs_fx['Legal Entity'].isnull()]['Currency'].to_list()


# In[7]:


#create a copy of raw data to make changes
staff_corr = staff_raw_data.copy()


# In[8]:


#create a dataframe containing adjustable currencies wo LE filter
staff_fx_part_1 = staff_corr[staff_corr['Currency'].isin(list_fx_wo_le)]
#create a dataframe containing adjustable currencies with LE filter
staff_fx_part_2 = staff_corr.merge(list_fx_le, on =['Currency', 'Legal Entity'], how='inner')


# In[9]:


#create a dataframe containing all adjustable currencies
staff_fx = pd.concat([staff_fx_part_1, staff_fx_part_2])


# In[10]:


#create a list of columns in a ranges:
hours_columns = list(range(inputs_ranges.loc[inputs_ranges['Range']=='Hours']['First column'].values[0] - 1, 
                      inputs_ranges.loc[inputs_ranges['Range']=='Hours']['Last column'].values[0]))

original_curr_columns = list(range(
    inputs_ranges.loc[inputs_ranges['Range']=='Original currency values']['First column'].values[0] - 1, 
                      inputs_ranges.loc[inputs_ranges['Range']=='Original currency values']['Last column'].values[0]))

usd_curr_columns = list(range(
    inputs_ranges.loc[inputs_ranges['Range']=='USD currency values']['First column'].values[0] - 1, 
                      inputs_ranges.loc[inputs_ranges['Range']=='USD currency values']['Last column'].values[0]))


# In[11]:


#create a dataframe with a reversal numbers 
staff_fx_reversal = staff_fx.copy()
staff_fx_reversal.iloc[:, hours_columns + original_curr_columns + usd_curr_columns] = staff_fx.iloc[:,
                                            hours_columns + original_curr_columns + usd_curr_columns] * -1


# In[12]:


#Create a dataframe containing adjusted currencies (origianl in USD)
staff_fx_adj = staff_fx.copy()
staff_fx_adj.iloc[:, original_curr_columns] = staff_fx_adj.iloc[:, usd_curr_columns]
staff_fx_adj['Currency'] = 'USD'
staff_fx_adj['FX Rate'] = 1


# In[13]:


#consolidate dataframes
staff_fx_out = pd.concat([staff_corr, staff_fx_reversal, staff_fx_adj])


# In[17]:


#save to output file to excel
staff_fx_out.to_excel(directory + 'out_staff_wo_iwo.xlsx', index=False)


# ### Check

# In[18]:


#Calculate sum of USD amount before and after correction
Amount_USD_before = round(staff_raw_data['Total Salaries_USD'].sum(), 1)
Amount_USD_after = round(staff_fx_out['Total Salaries_USD'].sum(), 1)
#Calculate sum of work hours amount before and after correction
Amount_Hours_before = round(staff_raw_data['Work hours'].sum(), 1)
Amount_Hours_after = round(staff_fx_out['Work hours'].sum(), 1)
#Calculate how many new rows created in out dataframe compare to raw_data df
New_rows = staff_fx_out.shape[0] - staff_raw_data.shape[0]


# In[19]:


window = Tk() 
#create a title
window.title('Result')
#set a size of the window and the place on the screen
window.geometry("300x200+500+300") 
#creating labels
label_1 = Label(window, text='Amount USD before: '+str(Amount_USD_before)).place(x = 30, y = 10) 
label_2 = Label(window, text='Amount USD after: '+str(Amount_USD_after)).place(x = 30, y = 30) 
label_3 = Label(window, text='Difference: '+str(Amount_USD_after-Amount_USD_before)).place(x = 30, y = 50) 

label_4 = Label(window, text='Amount of Work hours before: '+str(Amount_Hours_before)).place(x = 30, y = 80) 
label_5 = Label(window, text='Amount of Work hours after: '+str(Amount_Hours_after)).place(x = 30, y = 100) 
label_6 = Label(window, text='Difference: '+str(Amount_Hours_after-Amount_Hours_before)).place(x = 30, y = 120) 

label_7 = Label(window, text='New rows have been created: '+str(New_rows)).place(x = 30, y = 150) 

window.mainloop() 

