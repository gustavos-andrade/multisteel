# -*- coding: utf-8 -*-
"""
Created on Wed May  7 13:08:34 2025

@author: gustavo.andrade
"""

import pandas as pd
from pypdf import PdfReader

path = r'Z:\PUBLICO\GUSTAVO\29341 - brayton.pdf'

save = path.replace('.pdf', '.xlsx')

pdf = PdfReader(path)

list_tables = []

for i in range(len(pdf.pages)-1):
    
    text = pdf.pages[i].extract_text()

    item = text[text.find('(s)')+3:text.find('848420006,50')-1]

    un_amount = text[text.find('ITEM')+5:text.find('%')-1]

    total_amount = text[text.find('IPI')+3:text.find('NCMPRAZO')]

    qty = text[text.find('TOTAL')+6:text.find('Pç')]

    code = text[text.find('Pç')+3:text.find('75 Dia')]

    desc = text[text.find('DESCRIÇÃO')+10:text.find('Brayton')-3]

    pump = text[text.find('MODELO:')+8:text.find('TAG:')]

    api = text[text.find('API:')+5:text.find('FLUIDO')]
    
    objects = [item, code, desc, qty, pump, api, un_amount, total_amount]
    
    list_tables.append(objects)
    
seal = pd.DataFrame(list_tables, columns= ['ITEM', 'PRODUTO', 'DESCRIÇÃO', 'QTDE', 'BOMBA', 'PLANO API', 'VLR UNIT', 'VLR TOTAL'])

seal.to_excel(save, index = False)
