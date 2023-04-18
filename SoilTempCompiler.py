# -*- coding: utf-8 -*-
"""
Created on Thu Jul 28 13:37:53 2022

@author: benjamin.mccarthy
"""
import pandas,os, tkinter as tk, tkinter.simpledialog
from tkinter import filedialog

''' This will be the Working directory for the script. Any time it needs to be updated, just drop the 
Files in the Analytical folder and the script will automaticall finbd what it's looking for '''
root = tk.Tk()
root.withdraw()

file_path = filedialog.askdirectory()
Month = tk.simpledialog.askstring("Month", 'Please enter full month name')
Year = tk.simpledialog.askstring("year", 'Please enter full Year')


save_path = filedialog.askdirectory()
SaveName = save_path+str(Year)+str(Month)+'.xlsx'
print(file_path)



#%% Here we walk down our folder, pull the information regarding the temperature, rename the column and append it to our dataframe. 
# next we grab the details and append those to a dataframe. 
LastFile = True
for root, dirs, files in os.walk(file_path):
    for file in files:
        if(file.endswith("(Data EDT).xlsx")):
            dfTemp = pandas.read_excel(os.path.join(root,file))
            dfTemp['Date-Time (EST/EDT)'] = dfTemp['Date-Time (EST/EDT)'].replace('53','00')          
            dfTemp[str(file)[:6]] = dfTemp['Ch: 1 - Temperature   (°C)']
            dfTemp = dfTemp.drop(['#','Ch: 1 - Temperature   (°C)'], axis=1)
            dfInt = pandas.read_excel(os.path.join(root,file),sheet_name=2)
            print(os.path.join(root,file))
        
            if LastFile == True:
                dfData = dfTemp.copy()
                dfInfo = dfInt.copy()
                LastFile = False
            else:
                dfData = dfData.join(dfTemp.set_index('Date-Time (EST/EDT)'), on='Date-Time (EST/EDT)')
                dfInfo = dfInfo.append(dfInt)

dfs = {'Temp Results': dfData, 'Device Info': dfInfo}

writer = pandas.ExcelWriter(SaveName, engine='xlsxwriter')
for sheetname, df in dfs.items():  # loop through `dict` of dataframes
    df.to_excel(writer, sheet_name=sheetname)  # send df to writer
    worksheet = writer.sheets[sheetname]  # pull worksheet object
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        if isinstance(series.name, tuple):
            max_len = max(series.astype(str).map(len).max(),(len(series.name[1]))) + 1
        else:
            max_len = max(series.astype(str).map(len).max(),(len(series.name))) + 1   
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.save()
writer.close()
print(str(SaveName)+' saved to folder')

            
dfTemp = pandas.read_excel(r'Save location.xlsx')
