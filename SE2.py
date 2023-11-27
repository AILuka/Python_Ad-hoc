#!/usr/bin/env python
# coding: utf-8

# In[1]:


print('Loading...')


# In[2]:


import pandas as pd
import numpy as np
from tkinter import *
import tkinter as ttk
import os


# ### Inputs window

# In[3]:


#get data from entry field, workbook name
def show_message():
    global directory, workbook_drill_1_name, workbook_drill_2_name, workbook_iwo_report_name
    directory = entry_1.get()
    workbook_drill_1_name = entry_2.get()
    workbook_drill_2_name = entry_3.get()
    workbook_iwo_report_name = entry_4.get()
    root.destroy()

# create a window with inputs
root = Tk()
root.title("Inputs")
root.geometry("500x260+500+300") 

##parametrs
p_w = 80
p_x = 10
p_y = 10

###

lbl_1 = Label(root, text='Enter a file directory:')
lbl_1.place(x = p_x, y = p_y * 1) 

entry_1 = ttk.Entry(width=p_w)
entry_1.place(x = p_x, y = p_y * 3) 
#add sub folder IWO Recharges
entry_1.insert(0, os.getcwd() + '\IWO Recharges') #defauld entry text

###

lbl_2 = Label(root, text='Drill file_1 name:')
lbl_2.place(x = p_x, y = p_y * 6) 

entry_2 = ttk.Entry(width=p_w)
entry_2.place(x = p_x, y = p_y * 8)
entry_2.insert(0, 'in_Drill Margin to Fact_1.xlsx') #defauld entry text

###

lbl_3 = Label(root, text='Drill file_2 name:')
lbl_3.place(x = p_x, y = p_y * 11) 

entry_3 = ttk.Entry(width=p_w)
entry_3.place(x = p_x, y = p_y * 13)
entry_3.insert(0, 'in_Drill Margin to Fact_2.xlsx') #defauld entry text

###

lbl_4 = Label(root, text='Luxoft Staff in IWO file name:')
lbl_4.place(x = p_x, y = p_y * 16) 

entry_4 = ttk.Entry(width=p_w)
entry_4.place(x = p_x, y = p_y * 18)
entry_4.insert(0, 'in_Luxoft Staff in IWO.xlsx') #defauld entry text

btn = ttk.Button(text="Start", command=show_message)
btn.place(x = p_x, y = p_y * 21) 

root.mainloop()


# In[4]:


print('Doing the math, wait a minute...')


# ### Calculation

# In[6]:


#create a python directiry format
directory = directory.replace("\\","/")+'/'
#upload files
df_drill_1 = pd.read_excel(directory+workbook_drill_1_name, header=10)
df_drill_2 = pd.read_excel(directory+workbook_drill_2_name, header=10)
iwo_report = pd.read_excel(directory+workbook_iwo_report_name, header=0)

#upload inputs file
directory_inputs = os.getcwd().replace("\\","/")+'/'
#directory_inputs = '//Bull/Private$/SDC/Departments/FP&A/Python Fact/Staff expense/'

file_inputs = 'Inputs.xlsx'
inputs_region_guid = pd.read_excel(directory_inputs+file_inputs, sheet_name='Region_GUID', header=0)


# In[7]:


#combine drill dataframes
df_drill = pd.concat([df_drill_1, df_drill_2], ignore_index=True)
#drop subtotal "Total Numbering" (using not null locations)
df_drill = df_drill[df_drill['Location'].notnull()]
#rename columns with empty titles
df_drill = df_drill.rename(columns={'Unnamed: 0':'FBU_drill', 'Unnamed: 1':'Month', 'Unnamed: 2':'PIN'})

#fill all NaN values in columns with 0, in order to properly create pivot
df_drill = df_drill.fillna(0)


# In[8]:


#create a pivot from cognos drill
drill_pivot = df_drill.groupby(
    ['Source of Data', 'FBU_drill','Fin Project', 'Location', 'Luxoft LE', 'Currency'], as_index = False).agg(
    {'FX Rate':'mean', 'Amount USD Currency':'sum'})


# In[9]:


#combine drill and IWO report
iwo_1 = iwo_report.merge(
    drill_pivot, left_on='Invoice number', right_on='Source of Data', how='left')


# In[10]:


iwo_2 = iwo_1.copy()
#drop column from drill table
iwo_2 = iwo_2.drop(['Amount USD Currency'], axis='columns')
#fill all NaN values in numerical columns with 0, in order to properly calculate columns sum
iwo_2.loc[:, iwo_2.select_dtypes(include=[np.number]).columns] = iwo_2.select_dtypes(include=[np.number]).fillna(0)

#add columns
iwo_2['Hours'] = iwo_2['Hours charged'] * (-1)
iwo_2['Amount_LC'] = iwo_2['Total'] * (-1)
iwo_2['Amount_USD'] = iwo_2['Amount_LC'] / iwo_2['FX Rate']

iwo_2['Amount_LC_Other'] = iwo_2['Other costs'] * (-1)
iwo_2['Amount_USD_Other'] = iwo_2['Amount_LC_Other'] / iwo_2['FX Rate']

iwo_2['Total_Amount_LC'] = iwo_2['Amount_LC'] + iwo_2['Amount_LC_Other']
iwo_2['Total_Amount_USD'] = iwo_2['Amount_USD'] + iwo_2['Amount_USD_Other']

#add region GUID
iwo_2 = iwo_2.merge(inputs_region_guid, 
                              left_on='Location', right_on='Region', how='left')


# In[13]:


#check amount by FBU
iwo_check = iwo_2.groupby(
    ['Invoice number'], as_index = False).agg(
    {'Total_Amount_USD':'sum'})
#add values from drill pivot
iwo_check_2 = iwo_check.merge(drill_pivot[['Source of Data', 'Amount USD Currency']], 
                              left_on='Invoice number', right_on='Source of Data', how='outer')
iwo_check_2 = iwo_check_2.fillna(0)

iwo_check_2['difference'] = iwo_check_2['Total_Amount_USD'] - iwo_check_2['Amount USD Currency']
iwo_check_2['difference'] = round(iwo_check_2['difference'], 0)
iwo_check_3 = iwo_check_2.loc[iwo_check_2['difference']!=0]


# In[14]:


#save to interim output IWO file to excel
iwo_2.to_excel(directory + 'out_check_IWO.xlsx', index=False)
#save to interim output drill file to excel
drill_pivot.to_excel(directory + 'out_check_Drill.xlsx', index=False)
#save the table with USD FBU difference (potential errors)
iwo_check_2.to_excel(directory + 'out_check_amount.xlsx', index=False)


# In[15]:


#split cost to direct cost and other cost
iwo_3 = iwo_2[['Region GUID', 'Location', 'Luxoft LE', 'FBU', 'Fin project', 
                                       'PIN', 'Employee', 'Hours', 'Currency', 'FX Rate', 
                                       'IWO number', 'Amount_LC', 'Amount_LC_Other', 'Amount_USD', 'Amount_USD_Other']]
#create df with a direct cost
iwo_3_direct_cost = (iwo_3.loc[iwo_2['Amount_LC']!=0]
                     .drop(['Amount_LC_Other', 'Amount_USD_Other'], axis='columns')
                     .rename(columns={'Amount_LC':'Expense recharge from DXC', 
                                      'Amount_USD':'Expense recharge from DXC_USD'}))
#create df with an other cost
iwo_3_other_cost = (iwo_3.loc[iwo_2['Amount_LC_Other']!=0]
                    .drop(['Amount_LC', 'Amount_USD'], axis='columns')
                    .rename(columns={'Amount_LC_Other':'Expense recharge from DXC', 
                                      'Amount_USD_Other':'Expense recharge from DXC_USD'}))
iwo_3_other_cost.loc[:, ['Hours']] = 0
#combine tables of direct cost and other cost
iwo_4 = pd.concat([iwo_3_direct_cost, iwo_3_other_cost], ignore_index=True)
#rename columns
iwo_4 = iwo_4.rename(columns={'Luxoft LE':'Legal Entity', 'FBU':'Business Unit', 'Fin project':'Fin Project', 
                             'Employee':'Employee / Company', 'Hours':'Work hours', 
                              'Location':'Region', 'IWO number':'Request number'})
#add columns
iwo_4.insert(loc=0, column='Numbering', value=np.arange(1, len(iwo_4)+1)) #numbering
iwo_4.loc[:, ['Classification']] = 'PL'
iwo_4.loc[:, ['Total hours']] = iwo_4['Work hours']


# In[16]:


header_df = pd.read_excel(directory_inputs+'out_staff_wo_iwo.xlsx', nrows=0, header=0)
#header_df = pd.read_excel(directory_inputs+'out_staff_wo_iwo.xlsx', nrows=0, header=0)
iwo_out = pd.concat([header_df, iwo_4])


# In[17]:


iwo_out.to_excel(directory + 'out_iwo_recharges.xlsx', index=False)


# ### Check

# In[18]:


#calculate the nubmer of rows with empty FBU/invoice after merging
empty_FBU_check = iwo_1['Source of Data'].isnull().sum()
usd_FBU_check = len(iwo_check_3)


# In[19]:


window = Tk() 
#create a title
window.title('Result')
#set a size of the window and the place on the screen
window.geometry("350x150+500+300") 

#creating labels
#label_1 = Label(window, text='Number of empty FBU names: '+str(empty_FBU_check)).place(x = 30, y = 20) 

label_2 = Label(window, text='Number of Invoices with amount USD difference: '+str(usd_FBU_check)).place(x = 30, y = 50) 

window.mainloop() 


# In[ ]:




