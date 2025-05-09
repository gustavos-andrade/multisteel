# -*- coding: utf-8 -*-
"""
Created on Thu Sep 26 07:50:25 2024

@author: gustavo.andrade
"""

import pandas as pd
import os
import glob
from xls2xlsx import XLS2XLSX

path = (r'Z:/COMERCIAL/PEDIDOS DE BOMBAS - OSB/OSB/OSB/OSB-2017')
    
cot_folders = os.listdir(path)

cot_folders = cot_folders[:len(cot_folders) - 3]

xls = []

for j in range(len(cot_folders)):
    
    path_file = path + "/" + cot_folders[j]
    
    files = glob.glob(os.path.join(path_file, "*.xls"))
    
    if (len(files) == 0):
        pass
    else:
        xls.append(files)
    
    pdf_df = pd.DataFrame(xls)
    
list_xls = []

for row in range(len(pdf_df)):
    for col in range(2):
        if pdf_df.iloc[row,col] != None :
            list_xls.append(pdf_df.iloc[row,col])

    
for i in range(len(list_xls)):
    x2x = XLS2XLSX(list_xls[i])
    save_name = list_xls[i].replace('xls','xlsx')
    x2x.to_xlsx(save_name)
    
    