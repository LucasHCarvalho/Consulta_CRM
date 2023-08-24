import time
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By

options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)    

driver.get("https://guiamedico.cremesp.org.br/")

crm = 101213
text = 'crm'

driver.find_element(By.NAME, 'crm').send_keys(crm)
driver.find_element(By.XPATH, '/html/body/app-root/div/app-principal/div/div[1]/app-pesquisa-form/div[1]/div/form/div/div[2]/button').click()

time.sleep(1)

try:
    text = driver.find_element(By.XPATH, '//*[@id="DataTables_Table_0"]/tbody/tr/td[2]').text
except:
    text = 'CRM não encontrado'

with open('crm.csv', 'w', newline='', encoding='utf-8') as arquivo_csv:
    escritor_csv = csv.writer(arquivo_csv)
    escritor_csv.writerow([crm, text])

arquivo_csv.close()
