# -*- coding: utf-8 -*-
"""
Created on Thu Sep 19 08:21:38 2024

@author: gustavo.andrade
"""

import pandas as pd
import os
import glob
import re


path = (r'Z:/COMERCIAL/PEDIDOS DE BOMBAS - OSB/OSB/OSB/OSB-2019')
    
cot_folders = os.listdir(path)

cot_folders = cot_folders[:len(cot_folders) - 4]

data_raw = []

for j in range(len(cot_folders)):
    path_file = path + "/" + cot_folders[j]
    
    files = glob.glob(os.path.join(path_file, "*.xlsx"))
    
    pdf = glob.glob(os.path.join(path_file, "*.pdf"))
    
    forb = ["TÉCNICA", "TÉC", " TECNICA", "CÓPIA", 'AGENDA', 'GERADOR', 'CÓPIA','~$BB24', 'RZN', 'COTAÇÃO', 'FOLHA DE DADOS', 'PLANILHA', 'ESPECIFICAÇÃO', 'FD', '~$', 'RELAÇÃO']
    
    
    if (len(files) == 0):
        pass
    else:

        last = files[len(files)-1]
        
        ln = 0
            
        while (any(x in last.split("\\")[-1].upper() for x in forb)) == True:
            last = files[(len(files)-1)-ln]
            ln = ln + 1

        
        df = pd.read_excel(last, header=None)
    
        df = df.dropna(axis = 1, how = 'all', ignore_index = True)
        
        df = df.dropna(axis = 0, how = 'all', ignore_index = True)
        
        df.iloc[:,0] = df.iloc[:,0].str.upper()
        
        df.iloc[:,0] = df.iloc[:,0].str.replace("+","/")
        
        osb = df.iloc[4:11,:].dropna(axis = 1, how = 'all', ignore_index = True)
        
        osb = osb.dropna(axis = 0, how = 'all', ignore_index = True)
        
        tec_data = df[df.eq("FLUÍDO:    ").any(axis=1)].dropna(axis = 1, how = 'all')
        
        head = df.iloc[0:7,].dropna(axis = 1, how = 'all', ignore_index = True)
        
        head.columns = range(head.columns.size)
        
        date_index = head.isin(['DATA:']).any(axis=0).idxmax()+1
    
        date = head.iloc[0,date_index]
        
        cot_index = head.isin(['COTAÇÃO:']).any(axis=0).idxmax()+1
    
        cotation = head.iloc[2, cot_index]
        
        if (osb.iloc[0,0].count('.') > 1 or osb.iloc[0,0].count('/') > 1):
            customer = osb.iloc[1,2]
        
            contact = osb.iloc[3,2]
        
            phone = osb.iloc[4,2]
        
            mail = osb.iloc[4,2]
        
            req = osb.iloc[1,0].upper().replace("ORDEM DE COMPRA: ", "")
            
            freight = osb.iloc[4,0].upper().replace("TIPO DE FRETE:  ", "")
            
            issue_date =  osb.iloc[0,0][osb.iloc[0,0].find('-')+2:].replace(".","/")
            
            delivery_date = osb.iloc[1,0].upper().replace("DATA DE ENTREGA:  ", "")
            
            cnpj = osb.iloc[2,2]
            
            payment = osb.iloc[3,0]
            
            osb_number = osb.iloc[0,0].upper().replace("ORDEM DE FABRICAÇÃO: ", "")
        else:
            customer = osb.iloc[1,2]
        
            contact = osb.iloc[3,2]
        
            phone = osb.iloc[4,2]
        
            mail = osb.iloc[5,2]
        
            req = osb.iloc[1,0].upper().replace("PEDIDO DE COMPRA: ", "")
            
            freight = osb.iloc[4,0].upper().replace("TIPO DE FRETE:  ", "")
            
            issue_date =  osb.iloc[2,0].upper().replace("DATA DE EMISSÃO:  ", "")
            
            issue_date = issue_date.upper().replace("DATA DE EMISSÃO: ", "")
            
            delivery_date = osb.iloc[3,0].upper().replace("DATA DE ENTREGA:  ", "")
            
            cnpj = osb.iloc[2,2]
            
            payment = osb.iloc[5,0]
            
            osb_number = osb.iloc[0,0].upper().replace("ORDEM DE FABRICAÇÃO: ", "")          
    

        
    #    set_price = df[df.eq("PREÇO TOTAL, 01 CONJUNTO:|PRECIO TOTAL, 01 BOMBA:").any(axis=1)].dropna(axis = 1, how = 'all', ignore_index = True)
        
        check = df.iloc[:,0].str.contains("PREÇO TOTAL, 01 CONJUNTO:|PREÇO TOTAL, 01 BOMBA:|PRECIO TOTAL, 01 BOMBA:|PRECIO TOTAL, 01 CONJUNTO:|BOMBA CENTRIFUGA / BASE / ACOPLAMENTO / SELO MECÂNICO|BOMBA CENTRIFUGA / BASE / ACOPLAMENTO / SELO MECÂNICO|PREÇO TOTAL, 01 BOMBA CENTRÍFUGA:|PREÇO TOTAL, 01 CONJUNTO MOTOBOMBA CENTRÍFUGA:|PRECIO TOTAL, 01 CONJUNTO MOTOBOMBA:|BOMBA CENTRÍFUGA / BASE / ACOPLAMENTO|BOMBA CENTRIFUGA / SELO MECÂNICO / ACOPLAMENTO / BASE|PREÇO TOTAL, 01 BOMBA|PREÇO TOTAL, 01 CONJUNTO|PREÇO TOTAL: 01 CONJUNTO MOTO BOMBA:")
        
        df_check = pd.DataFrame(check)
    
        price_loc = df_check[df_check[0]==True].index.values.astype(int)
        
        ptype_check = df.iloc[:,0].str.contains("CENTRÍFUGA ")
        
        ptype_check = pd.DataFrame(ptype_check)
        
        ptype_loc = ptype_check[ptype_check[0]==True].index.values.astype(int)
        
        itype_check = df.iloc[:,0].str.contains("SERIE:|SÉRIE:")
        
        itype_check = pd.DataFrame(itype_check)
        
        itype_loc = ptype_check[itype_check[0]==True].index.values.astype(int)

        for i in range(len(tec_data)):
            
            item_number = i + 1
            
            tec_inf = df.iloc[tec_data.index[i]:tec_data.index[i]+23,].dropna(axis = 1, how = 'all', ignore_index = True)
    
            tec_inf.columns = range(tec_inf.columns.size)
            
            # ptype = df.iloc[ptype_loc[i],0][41:51]
            
            ptype =  df.iloc[ptype_loc[i],0][df.iloc[ptype_loc[i],0].find('CENTRÍFUGA ')+11:df.iloc[ptype_loc[i],0].find(' - MULTISTEEL')]

            imp_type =  df.iloc[itype_loc[i],0][df.iloc[itype_loc[i],0].find('/')-8:df.iloc[itype_loc[i],0].find('/')]
            
            imp_type = re.sub("[0-9 : / E]","",imp_type)
            
            # imp_type = df.iloc[9,0][56:60].strip()
    
            pump_type = df.iloc[itype_loc[i],0][df.iloc[itype_loc[i],0].find('/')-4:df.iloc[itype_loc[i],0].find('/')+5]
    
            pump_type = re.sub("[^0-9/]","",pump_type)
            
            qty = df.iloc[itype_loc[i],0][df.iloc[itype_loc[i],0].find(':')+1:df.iloc[itype_loc[i],0].find(':')+7]
            
            qty = int(re.sub("[^0-9/]","", qty))
            
            fluid = tec_inf.iloc[0, 1]
    
            temp = tec_inf.iloc[1, 1]
    
            density = tec_inf.iloc[2, 1]
    
            visc = tec_inf.iloc[3, 1]
    
            flow = tec_inf.iloc[4, 1]
    
            hman = tec_inf.iloc[5, 1]
    
            dsimp = tec_inf.iloc[6, 1]
    
            dmin = tec_inf.iloc[7, 1]
    
            dmax = tec_inf.iloc[8, 1]
    
            seal = tec_inf.iloc[9, 1]
    
            npshreq = tec_inf.iloc[10, 1]
    
            npshdisp = tec_inf.iloc[11, 1]
    
            dsuc = tec_inf.iloc[0, 3]
    
            ddis = tec_inf.iloc[1, 3]
    
            efficiency = tec_inf.iloc[2, 3]
    
            power = tec_inf.iloc[3, 3]
    
            motor = tec_inf.iloc[4, 3]
    
            carcass_mat = tec_inf.iloc[5, 3]
    
            imp_material = tec_inf.iloc[6, 3]
    
            cover_material = tec_inf.iloc[7, 3]
    
            wear_ring = tec_inf.iloc[8, 3]
    
            glove_material = tec_inf.iloc[9, 3]
    
            axis_material = tec_inf.iloc[10, 3]
            
            addinf = ""
            
            for line in range(13,len(tec_inf)):
                addinf = addinf + str(tec_inf.iloc[line, 0])
    
            # addinf = str(tec_inf.iloc[13, 0]) + str(tec_inf.iloc[14, 0]) + str(tec_inf.iloc[15, 0]) + str(tec_inf.iloc[16, 0]) + str(tec_inf.iloc[17, 0]) + str(tec_inf.iloc[18, 0]) + str(tec_inf.iloc[19, 0]) + str(tec_inf.iloc[20, 0])
    
            fac_column = tec_inf.isin(['FATOR']).any(axis=0).idxmax()
    
            pump_price = "R$ " + str(tec_inf.iloc[1, fac_column-1]) + " - FATOR " + str(tec_inf.iloc[1, fac_column])
    
            base = "R$ " + str(tec_inf.iloc[2, fac_column-1]) + " - FATOR " + str(tec_inf.iloc[2, fac_column])
    
            coupling = "R$ " + str(tec_inf.iloc[3, fac_column-1]) + " - FATOR " + str(tec_inf.iloc[3, fac_column])
    
            seal_price = "R$ " + str(tec_inf.iloc[4, fac_column-1]) + " - FATOR " + str(tec_inf.iloc[4, fac_column])
    
            motor_price = "R$ " + str(tec_inf.iloc[5, fac_column-1]) + " - FATOR " + str(tec_inf.iloc[5, fac_column])
    
            add1 = "R$ " + str(tec_inf.iloc[6, fac_column-1]) + " - FATOR " + str(tec_inf.iloc[6, fac_column])
    
            add2 = "R$ " + str(tec_inf.iloc[7, fac_column-1]) + " - FATOR " + str(tec_inf.iloc[7, fac_column])
    
            try:
                price = df.iloc[price_loc[i],3:9].dropna()
                
                price = float(price.iloc[0])
                
                total = price * qty

            except (IndexError):
                price = 0
           
            data_raw.append([cotation, osb_number, date, customer, cnpj, contact, phone, mail, req, issue_date, delivery_date, freight, payment, item_number, qty, ptype, imp_type, pump_type, fluid, temp, density, visc, flow, hman, dsimp, dmin, dmax, seal, npshreq, npshdisp, dsuc, ddis, efficiency, power, motor, carcass_mat, imp_material, cover_material, wear_ring, glove_material, axis_material, addinf, pump_price, base, coupling, seal_price, motor_price, add1, add2, price, total])
                            
        
bbreg = pd.DataFrame(data_raw, columns= ['PRICING_NUMBER', 'OSB', 'DATE', 'CUSTOMER', 'CNPJ', 'CONTACT', 'PHONE', 'EMAIL', 'REQ','ISSUE DATE', 'DELIVERY DATE', 'FREIGHT', 'PAYMENT', 'ITEM_NUMBER', 'QTY', 'TYPE', 'MODEL', 'SIZE', 'FLUID', 'TEMPERATURE', 'DENSITY', 'VISCOSITY', 'FLOW', 'HEAD', 'D_SELECTED', 'D_MIN', 'D_MAX', 'SEAL', 'NPSH_REQ', 'NPSH_DISP', 'D_SUCTION', 'D_DISCHARGE', 'EFFICIENCY', 'POWER', 'MOTOR', 'CARC_MAT', 'IMP_MAT', 'COV_MAT', 'WEAR_RING', 'GLOVE', 'SHAFT', 'ADD_INF', 'PUMP_PRICE','BASE', 'COUPLING', 'SEAL_PRICE', 'MOTOR', 'ADD1', 'ADD2', 'PRICE', 'TOTAL PRICE'])

bbreg.to_excel(r'C:\Users\gustavo.andrade\Downloads\osbreg19.xlsx', index = False)

bbreg.to_csv(r'C:\Users\gustavo.andrade\Downloads\osbreg19.csv', index = False)





