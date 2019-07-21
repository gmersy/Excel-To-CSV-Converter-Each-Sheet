#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jun 13 22:46:38 2019

@author: gabemersy

Takes .xlsx file and returns a .csv file for each sheet within the excel file

Loops through a folder and converts all .xlsx files to csvs for each respective sheet

Libraries/packages that need to be installed: pip, wheel, pandas
"""

import pandas as pd

import os

# Converts each individual sheet to a seperate CSV. Sheet one is 0, sheet two is 1, sheet three is 2, etc.
def excelToCSV(fileName, path, excelFileObj):
    numChars = len(fileName)
    numNoExten = numChars - 5
    
    # Separating file name
    noExten = ''
    if numNoExten > 5:
        for i in range(0, numNoExten):
            noExten += fileName[i]
    else:
        print("File name needs to be greater than 5 characters")
    
    numSheets = len(excelFileObj.sheet_names)
    # Scanning through sheets sheet 0 is first element of sheets array
    for i in (0, numSheets):
        df = pd.read_excel(path, sheet_name = i, encoding='utf-8')

        # Removes empty rows to correct dimensions
        df = df.dropna(axis = 1, how = 'all')
        # Removes empty columns to correct dimensions
        df = df.dropna(axis = 0, how = 'all')
        
        # New file string
        newFileName = noExten + str(i) + ".csv"
        
        df.to_csv(newFileName, sep=',', encoding='utf-8', index=False)
        

# Make sure you are in the correct folder 
folder = '/Users/gabemersy/Desktop/' 

# Loop through each file in the folder code
for filename in os.listdir(folder):
    if filename.endswith('.xlsx'):
        
        # Run conversion 
        path = folder + filename
        excelFileObj = pd.ExcelFile(path)
        excelToCSV(filename, path, excelFileObj)
    else:
        continue
