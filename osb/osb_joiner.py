# -*- coding: utf-8 -*-
"""
Created on Tue Nov 26 14:49:34 2024

@author: gustavo.andrade
"""

import pandas as pd

osb20 =  pd.read_excel('C:/Users/gustavo.andrade/Downloads/osbreg20.xlsx')

osb21 =  pd.read_excel('C:/Users/gustavo.andrade/Downloads/osbreg21.xlsx')

osb22 =  pd.read_excel('C:/Users/gustavo.andrade/Downloads/osbreg22.xlsx')

osb23 =  pd.read_excel('C:/Users/gustavo.andrade/Downloads/osbreg23.xlsx')

osb24 =  pd.read_excel('C:/Users/gustavo.andrade/Downloads/osbreg24.xlsx')

osb25 =  pd.read_excel('C:/Users/gustavo.andrade/Downloads/osbreg25.xlsx')

osb_df = [osb20, osb21, osb22, osb23, osb24, osb25]

osbreg = pd.concat(osb_df)

osbreg.to_excel(r'C:\Users\gustavo.andrade\Downloads\OSBreg.xlsx', index = False)


osb20 =  pd.read_csv('C:/Users/gustavo.andrade/Downloads/osbreg20.csv')

osb21 =  pd.read_csv('C:/Users/gustavo.andrade/Downloads/osbreg21.csv')

osb22 =  pd.read_csv('C:/Users/gustavo.andrade/Downloads/osbreg22.csv')

osb23 =  pd.read_csv('C:/Users/gustavo.andrade/Downloads/osbreg23.csv')

osb24 =  pd.read_csv('C:/Users/gustavo.andrade/Downloads/osbreg24.csv')

osb25 =  pd.read_csv('C:/Users/gustavo.andrade/Downloads/osbreg25.csv')

osb_df = [osb20, osb21, osb22, osb23, osb24, osb25]

osbreg = pd.concat(osb_df)

osbreg.to_csv(r'C:\Users\gustavo.andrade\Downloads\OSBreg.csv', index = False)
