# Importing libraries 
 
import pandas as pd
import os
fileName = '2022-0000 Sample PM Tool.xlsx'
filePath = r'C:\Users\mary\Downloads'
file=os.path.join(filePath,fileName)
df=pd.read_excel(file)
df
to_drop = ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3','Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11', 'Unnamed: 13', 'Unnamed: 14', 'Unnamed: 16', 'Unnamed: 17', 'Unnamed: 19', 'Unnamed: 20', 'Unnamed: 21', 'Unnamed: 22', 'Unnamed: 23', 'Unnamed: 24', 'Unnamed: 28', 'Unnamed: 30', 'Unnamed: 32', 'Unnamed: 33', 'Unnamed: 34', 'Unnamed: 35', 'Unnamed: 36', 'Unnamed: 37', 'Unnamed: 39', 'Unnamed: 40','Unnamed: 41']
df.drop(columns=to_drop, inplace=True)
print(df)
