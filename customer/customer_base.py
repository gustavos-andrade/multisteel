# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 07:44:19 2025

@author: gustavo.andrade
"""
import pandas as pd
import numpy as np
import os
import glob

customer_base = pd.read_excel(r'C:\Users\gustavo.andrade\Documents\BASE CLIENTES\Dados Relatório de Empresas.xls')

# customer_base = customer_base.drop(customer_base.iloc[29307:,].index)

customer = customer_base.dropna(how = 'all')

customer = customer.dropna(axis = 1, how = 'all')

customer = customer.iloc[9:,].reset_index(drop=True)

customer['test'] = customer.iloc[:,0].str.contains("MULTISTEEL COMERCIO DE BOMBAS ESPECIAIS LTDA                ")

for multi in range(len(customer)):
    if customer.iloc[multi,36] == True:
        customer.iloc[multi:multi+10, 36] = 'delete'
        
customer['test'] = customer['test'].astype(str)

customer = customer[customer["test"].str.contains("delete") == False]

customer = customer.drop('test', axis = 'columns')

customer['test'] = customer.iloc[:,0].str.contains("Razão Social Completa")

info = []

for i in range(len(customer)):
    if customer.iloc[i,36] == True:
        info.append(customer.iloc[i-1:i+9,0:34].reset_index(drop = True))

# line = []

# for j in range(len(info)):
#     temp = []
#     for lin in range(len(info[j])):
#         for col in range(len(info[j].columns)):
#             if (pd.isna(info[j].iloc[lin, col]) == False):
#                 temp.append(info[j].iloc[lin, col])
#     line.append(pd.DataFrame(temp).T)
    

temp = []

for j in range(len(info)):
    info[j] = info[j][info[j].iloc[:,0].str.contains('Complemento|Home|Contato') == False].reset_index(drop = True)
    code = info[j].iloc[0, 0]
    name1 = info[j].iloc[0,12]
    name2 = info[j].iloc[1,6]
    cnpj = info[j].iloc[0,21]
    state_enroll = info[j].iloc[0,29]
    address = info[j].iloc[2,0]
    phone = info[j].iloc[2,24]
    zip_code = info[j].iloc[3,0]
    city = info[j].iloc[3,5]
    street = info[j].iloc[2,18]
    state = info[j].iloc[3,11]
    date_reg = info[j].iloc[3,32]
    date_req = info[j].iloc[4,32]
    std_mail = info[j].iloc[4,2]
    danfe_mail = info[j].iloc[5,2]       
    temp.append([code, name1, name2, cnpj, state_enroll, address, street, city, state, zip_code, std_mail, danfe_mail, date_reg, date_req])
            
df = pd.DataFrame(temp, columns= ['CÓDIGO', 'NOME', 'RAZAO SOCIAL', 'CNPJ', 'INS EST', 'ENDERECO', 'BAIRRO', 'CIDADE', 'UF', 'CEP', 'EMAIL PADRAO','EMAIL DANFE', 'DATA REG', 'ULT COMP'])

df = df[df.iloc[:,0].str.contains('Razão') == False].reset_index(drop = True)

df['CNPJ'] = df['CNPJ'].str.replace("CNPJ:","", regex= True)

df.to_excel(r'C:\Users\gustavo.andrade\Downloads\Base_Clientes.xlsx', index = False, sheet_name = 'Customer')

# customer = customer[customer.eq("FLUÍDO:    ").any(axis=1)].dropna(axis = 1, how = 'all')
#customer = customer.apply(lambda x: pd.Series(x.dropna().values))