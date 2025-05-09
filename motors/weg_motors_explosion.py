# -*- coding: utf-8 -*-
"""
Created on Tue Sep  3 18:59:37 2024

@author: gusta
"""
import time
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import re

driver = webdriver.Firefox()


#Acessing WEG site
driver.get("https://www.weg.net/catalog/weg/BR/en/login")

# Login
username = driver.find_element(By.ID , "j_username")
username.clear()
username.send_keys("orcamentos3@bombasmultisteel.com.br")
password = driver.find_element(By.ID , "j_password")
password.clear()
password.send_keys("MULTI123@")
login_button = driver.find_element(By.XPATH , '//button[text()="Login"]')
login_button.click()

wait = WebDriverWait(driver, timeout=10)
cokkie = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[7]/div[1]/div/div/div[2]/a/div'))).click()
time.sleep(2)

action = ActionChains(driver)

# Select Menu items
products = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[2]/div/nav/div/div[4]/ul/li[2]/a'))).click()
electric_motors = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[2]/div/nav/div/div[4]/ul/li[2]/ul/li[1]/a'))).click()
tri_low_voltage = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[1]/div[1]/div/div/div[2]/div/div[2]/ul/li/a'))).click()
explosion_atm = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[1]/div[1]/div/div/div[2]/div/div[3]/ul/li/a'))).click()
explosion_proof = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[1]/div[1]/div/div[1]/div[2]/div/div[1]/ul/li/a'))).click()

w22 = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[1]/div[1]/div/div[1]/div[2]/div/div[1]/ul/li/a'))).click()
efficiency = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[1]/div[1]/h4/a'))).click()
ir3 = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[1]/div[2]/div/div[1]/ul/li[2]/form/label/input'))).click()
time.sleep(1)
sel_apply = driver.find_element(By.LINK_TEXT, "Aplicar seleção").click()
time.sleep(2)

# voltage = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[8]/div[1]/h4/a'))).click()
# voltage3 = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[8]/div[2]/div/div[1]/ul/li[2]/form/label'))).click()
# time.sleep(1)
# sel_apply = driver.find_element(By.LINK_TEXT, "Aplicar seleção").click()
# time.sleep(2)

foot = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[11]/div[1]/h4/a'))).click()
w_foot = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[11]/div[2]/div/div[1]/ul/li[1]/form/label'))).click()
time.sleep(1)
sel_apply = driver.find_element(By.LINK_TEXT, "Aplicar seleção").click()
time.sleep(2)

flange = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[12]/div[1]/h4/a'))).click()
wt_flange = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[12]/div[2]/div/div[1]/ul/li[3]/form/label'))).click()
time.sleep(1)
sel_apply = driver.find_element(By.LINK_TEXT, "Aplicar seleção").click()
time.sleep(2)


pg_total = driver.find_element(By.XPATH,'/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[3]/div[2]/div[3]/div[1]/ul').text

pages_total = pg_total[pg_total.find('...'):pg_total.find('Próximo')]

# pages_total = pg_total[len(pg_total)-15:len(pg_total)-7]

pages_total = int(re.sub("[^0-9]","",pages_total))

data = []

ActionChains(driver).scroll_by_amount(0, -300).perform()

for page in range(pages_total):
    
    checklist = driver.find_elements(By.LINK_TEXT, "Consultar")
    time.sleep(2)

    for i in range(len(checklist)):
        check = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Consultar"))).click()
        ActionChains(driver).scroll_by_amount(0, 30).perform()        
        # if i == (round(len(checklist)/2,0)-2):
        #     ActionChains(driver).scroll_by_amount(0, 700).perform()
        time.sleep(8)
        
        
    tbody = driver.find_element(By.XPATH, "/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[3]/div[2]/div[2]/table/tbody")


    for tr in tbody.find_elements(By.XPATH, '//tr'):
        row = [item.text for item in tr.find_elements(By.XPATH, './/td')]
        data.append(row)
        
    if (page == pages_total-1):
        pass
    else:
        next_pg = driver.find_element(By.LINK_TEXT, "Próximo").click()
        time.sleep(6)

driver = webdriver.Firefox()
        
#Acessing WEG site
driver.get("https://www.weg.net/catalog/weg/BR/en/login")

# Login
username = driver.find_element(By.ID , "j_username")
username.clear()
username.send_keys("orcamentos3@bombasmultisteel.com.br")
password = driver.find_element(By.ID , "j_password")
password.clear()
password.send_keys("MULTI123@")
login_button = driver.find_element(By.XPATH , '//button[text()="Login"]')
login_button.click()

wait = WebDriverWait(driver, timeout=10)
cokkie = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[7]/div[1]/div/div/div[2]/a/div'))).click()
time.sleep(2)

action = ActionChains(driver)

# Select Menu items
products = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[2]/div/nav/div/div[4]/ul/li[2]/a'))).click()
electric_motors = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[2]/div/nav/div/div[4]/ul/li[2]/ul/li[1]/a'))).click()
tri_low_voltage = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[1]/div[1]/div/div/div[2]/div/div[2]/ul/li/a'))).click()
explosion_atm = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[1]/div[1]/div/div/div[2]/div/div[3]/ul/li/a'))).click()
explosion_proof = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[1]/div[1]/div/div[1]/div[2]/div/div[1]/ul/li/a'))).click()

w21 = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[1]/div[1]/div/div[1]/div[2]/div/div[2]/ul/li/a'))).click()

# voltage = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[8]/div[1]/h4/a'))).click()
# voltage3 = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[8]/div[2]/div/div[1]/ul/li[2]/form/label'))).click()
# time.sleep(1)
# sel_apply = driver.find_element(By.LINK_TEXT, "Aplicar seleção").click()
# time.sleep(2)

foot = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[10]/div[1]/h4/a'))).click()
w_foot = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[10]/div[2]/div/div[1]/ul/li[1]/form/label/input'))).click()
time.sleep(1)
sel_apply = driver.find_element(By.LINK_TEXT, "Aplicar seleção").click()
time.sleep(2)

flange = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[12]/div[1]/h4/a'))).click()
wt_flange = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[2]/div/div[12]/div[2]/div/div[1]/ul/li[3]/form/label'))).click()
time.sleep(1)
sel_apply = driver.find_element(By.LINK_TEXT, "Aplicar seleção").click()
time.sleep(2)

pg_total = driver.find_element(By.XPATH,'/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[3]/div[2]/div[3]/div[1]/ul').text

pages_total = pg_total[pg_total.find('...'):pg_total.find('Próximo')]

# pages_total = pg_total[len(pg_total)-15:len(pg_total)-7]

pages_total = int(re.sub("[^0-9]","",pages_total))


ActionChains(driver).scroll_by_amount(0, -300).perform()

for page in range(pages_total):
    
    checklist = driver.find_elements(By.LINK_TEXT, "Consultar")
    time.sleep(2)

    for i in range(len(checklist)):
        check = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Consultar"))).click()
        ActionChains(driver).scroll_by_amount(0, 30).perform()        
        # if i == (round(len(checklist)/2,0)-2):
        #     ActionChains(driver).scroll_by_amount(0, 700).perform()
        time.sleep(8)
        
        
    tbody = driver.find_element(By.XPATH, "/html/body/div[4]/div[1]/div/div[2]/div[8]/div[2]/div/div[3]/div[2]/div[2]/table/tbody")


    for tr in tbody.find_elements(By.XPATH, '//tr'):
        row = [item.text for item in tr.find_elements(By.XPATH, './/td')]
        data.append(row)
        
    if (page == pages_total-1):
        pass
    else:
        next_pg = driver.find_element(By.LINK_TEXT, "Próximo").click()
        time.sleep(6)


        

df = pd.DataFrame.from_dict(data)

df = df[[2, 3, 5]]

df.replace('', np.nan, inplace=True)

motors = df.dropna()

motors[5] = motors[5].str.split('/').str[0]

motors = motors.reset_index(drop=True)

motors = motors.rename(columns={2:'DESCRIPTION', 3: 'CODE', 5:'PRICE'})

motors['BOX_POS'] = motors['DESCRIPTION'].str[-4:]

motors['COOLING'] = motors['DESCRIPTION'].str[-19:-6]

motors['COOLING'] = motors['COOLING'].str.strip()

# motors['FREQUENCY'] = motors['DESCRIPTION'].str[-25:-21]

motors['FREQUENCY'] = 60

motors['VOLTAGE'] = np.nan

motors['POWER'] = np.nan

motors['POLES'] = np.nan

motors['MODEL'] = np.nan

motors['EFFICIENCY'] = np.nan

for k in range(motors.shape[0]):
    motors.iloc[k, 6] = motors.iloc[k,0][motors.iloc[k,0].find('3F') + 2:motors.iloc[k, 0].find('3F') + 16]
    
k = 0

for k in range(motors.shape[0]):
    motors.iloc[k, 7] = motors.iloc[k,0][motors.iloc[k,0].find('cv') -6:motors.iloc[k, 0].find('cv') + 2]
    
k = 0

for k in range(motors.shape[0]):
    motors.iloc[k, 8] = motors.iloc[k,0][motors.iloc[k,0].find('cv') +3:motors.iloc[k, 0].find('cv') + 6]
    
k = 0

for k in range(motors.shape[0]):
    motors.iloc[k, 9] = motors.iloc[k,0][motors.iloc[k,0].find('cv') +6:motors.iloc[k, 0].find('cv') + 16]


motors['TYPE'] = motors['DESCRIPTION'].str[0:3]

motors['POWER'] = motors['POWER'].str.extract(r'(\d+[.\d]*)')

motors['POWER'] = motors['POWER'].astype(float)

motors['VOLTAGE'] = motors['VOLTAGE'].str.split(' V').str[0]

motors['POLES'] = motors['POLES'].str.strip()

motors['BOX_POS'] = motors['BOX_POS'].str.strip()

motors['MODEL'] = motors['MODEL'].str.replace('3F.*', '', regex=True)

motors['MODEL'] = motors['MODEL'].str.strip()

k = 0

for k in range(motors.shape[0]):
    motors.iloc[k, 10] = motors.iloc[k,0][motors.iloc[k,0].find('cv') -23:motors.iloc[k, 0].find('cv') -3]
   
    if (motors.iloc[k, 11] == 'W22'):
        motors.iloc[k, 10] = 'IR3 Premium'
        motors.iloc[k, 11] = 'W22Xdb'        
    else:
        motors.iloc[k, 10] = 'Explosão Standard'
    


# motors.to_csv(path_or_buf='C:/Users/gusta/Downloads/motors.csv')

# motors.to_csv(r'C:\Users\gustavo.andrade\Downloads\motors.csv', sep=';', encoding='utf-8')

motors.to_excel(r'C:\Users\gustavo.andrade\Downloads\motors_explosion_proof.xlsx', index = False)

driver.close()
