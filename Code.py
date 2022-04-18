# Importing libraries 
import numpy as np    
import pandas as pd
import xlrd
import xdrlib, sys
import os
import openpyxl
filename = '2022-0000 Sample PM Tool.xlsx'
filepath = r'c:\Users\mary\Downloads'
file=os.path.join(filepath,filename)
df=pd.read_excel(file)
df
