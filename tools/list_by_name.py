# -*- coding: utf-8 -*-
"""
Created on Fri Sep 27 07:11:10 2024

@author: gustavo.andrade
"""

import pandas as pd
import os
import glob
import re

path = (r'Z:/COMERCIAL/COTAÇÃO DE BOMBAS/Cotação de Bombas/2024')
    
cot_folders = os.listdir(path)

cot_folders = cot_folders[:len(cot_folders) - 3]

xls = []

for j in range(len(cot_folders)):
    
    path_file = path + "/" + cot_folders[j]
    
    prefixed = [filename for filename in os.listdir(path_file) if filename.startswith("~$")]
    
    if (len(prefixed) == 0):
        pass
    else:
        table = [prefixed, path_file]
        xls.append(table)
    
    pdf_df = pd.DataFrame(xls)