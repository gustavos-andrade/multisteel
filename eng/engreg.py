# -*- coding: utf-8 -*-
"""
Created on Fri Mar 28 07:15:42 2025

@author: gustavo.andrade
"""
import pandas as pd
import os

path = (r'Z:\ENGENHARIA\LISTA DE BOMBAS 10-2008')

exc_folders = ['Lista de reposição por O.S.B', 'Lista de Reposição de Bombas', 'LISTA DE BOMBA MODELO', 'INTERCÂMBILIDADE','BPDS/~$BPDS']


list_files = []

eng_list = []

for (dirpath, dirnames, filenames) in os.walk(path):
    for filename in filenames:
        if (filename.endswith((".xls", '.xlsx')) and (any(x in dirpath.split(os.sep) for x in exc_folders) == False)):
            file_path = dirpath + '/' + filename
            list_files.append(file_path)


list_files = [file for file in list_files if "~" not in file]
        

for i in range(len(list_files)):
    
    file = list_files[i]
    
    df1 = pd.read_excel(list_files[i], header = None, sheet_name = 0)
    
    df1 = df1.dropna(axis = 1, how = 'all', ignore_index = True)
    
    df1 = df1.dropna(axis = 0, how = 'all', ignore_index = True)
    
    df2 = pd.read_excel(list_files[i], header = None, sheet_name=1)
    
    df2 = df2.dropna(axis = 1, how = 'all', ignore_index = True)
    
    df2 = df2.dropna(axis = 0, how = 'all', ignore_index = True)
    
    df = pd.concat([df1,df2])
    
    try:
        date_check = df.iloc[:,1].str.contains("ELABORADO / DATA")
        
        date_check = pd.DataFrame(date_check)
        
        date_loc = date_check[date_check[1]==True].index.values.astype(int)
        
        date_index = date_loc[0] + 1
        
        date = df.iloc[date_index,1]
        
    except:
        
        date = df.iloc[1,8]
        
        df = df.drop(index = range(0,2)).reset_index(drop=True)
        
    
    pump = df.iloc[0,2]
    
    proj = df.iloc[1,2]
    
    customer = df.iloc[0,4]
    
    osb_set = df.iloc[2,2]
    
    osb = df.iloc[1,4]
    

    
    for j in range(4,len(df)):
        if (isinstance(df.iloc[j,0], int) == False):
                j += 1
        else:
            item = [osb, pump, customer, date, proj, osb_set, df.iloc[j,0], df.iloc[j,1], df.iloc[j,2], 
                    df.iloc[j,3], df.iloc[j,4], df.iloc[j,5], df.iloc[j,6], df.iloc[j,7], df.iloc[j,8]]
        eng_list.append(item)
        
engreg = pd.DataFrame(eng_list, columns= ['OSB', 'BOMBA', 'CLIENTE', 'DATA', 'PROJ', 'CONJ', 'POS', 'QTD', 'DENOMINACAO', 'DES N','DIMENSÕES', 'MATERIAL', 'PESO', 'MOD N', 'FORNEC/OSB'])
    
engreg.to_excel(r'C:\Users\gustavo.andrade\Downloads\ENGREG.xlsx', index = False, sheet_name = 'Engeneering')

engreg.to_csv(r'C:\Users\gustavo.andrade\Downloads\ENGREG.csv', index = False)
