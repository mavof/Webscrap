#!/usr/bin/env python
# coding: utf-8

# # Online Files Download

# In[3]:


#Regarding Browser and Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

#Manage Files
from pathlib import Path 
import glob
import os
import shutil

#Time Sleep control
import time

#Interacting with excel files
import pandas as pd

servico = Service(ChromeDriverManager().install())

#Functions to make the browser wait for elements
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#Avoiding Downloads Errors
downloads_path = str(Path.home() / "Downloads")
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": downloads_path,
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

#Open the Browser
navegador = webdriver.Chrome(service=servico,options=options)

#Read the information to download the files
tabela = pd.read_excel("Reference.xlsx")

#Run the routine for each line in the information source file
for linha in tabela.index:
    Ag = str(tabela.loc[linha,"AG"])
    CC = str(tabela.loc[linha,"CC"])
    Login = str(tabela.loc[linha,"Login"])
    Senha = str(tabela.loc[linha,"Password"])
    target = str(tabela.loc[linha,"Target"])
    
    #Open the website
    navegador.get("https://www.santander.com.br/empresas")
    time.sleep(2)

    #Access 
    navegador.find_element(By.NAME,
        'txtAgencia').send_keys(Ag)
    navegador.find_element(By.NAME,
        'txtConta').send_keys(CC)
    navegador.find_element(By.NAME,
        'txtConta').send_keys(Keys.ENTER)

    #Access
    time.sleep(15)
    navegador.find_element(By.ID,
        'formGeral:usuario').send_keys(Login)
    navegador.find_element(By.ID,
        'formGeral:senhaAcesso').send_keys(Senha)
    navegador.find_element(By.ID,
        'formGeral:senhaAcesso').send_keys(Keys.ENTER)

    #Closing the introduction Pou-up #Optimization Opportunity .By Method
    time.sleep(8)
    elemento = WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.ID,
        'formGeral:btnCloseModalLightBox')))
    time.sleep(1) # Avoiding errors
    elemento.click()

    #Click Extrato #Optimization Opportunity .By Method
    elemento = WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.XPATH,
        '/html/body/section/div/div/form/div[2]/div[2]/ul/li[2]/a')))
    time.sleep(1) 
    elemento.click()

    #Click Consultar
    elemento = WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.ID,
        'formGeral:j_id_8t:1:j_id_8y')))
    time.sleep(1) 
    elemento.click()

    #Select Time Period (Opening DropDown List)
    elemento = WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.ID,
        'extratoDatePicker')))
    time.sleep(1) 
    elemento.click()

    #Select Time Period (Mês atual [Current Month]) #Optimization Opportunity .By Method
    elemento = WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.XPATH,
        '/html/body/div[9]/div[1]/ul/li[5]')))
    time.sleep(1) 
    elemento.click()

    #Optimization Opportunity to find the botton of the page
    time.sleep(5)
    pyautogui.press("pgdn", presses=7)

    #Saving .PDF File
    time.sleep(5)
    navegador.find_element(By.ID,
        'formGeral:salvarPDF').click()

    #Closing Pou-up after download #Optimization Opportunity .By Method
    elemento = WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.XPATH,
        '/html/body/section/div/div/form/div[3]/div[8]/div[1]/a')))
    time.sleep(1) 
    elemento.click()

    #Finding the donwloaded file and saving in the target folder
    time.sleep(10)
    downloads_path = str(Path.home() / "Downloads")
    downloads_path = downloads_path + "\*"
    list_of_files = glob.glob(downloads_path) 
    latest_file = max(list_of_files, key=os.path.getctime)
    original = latest_file 
    shutil.move(original, target)

    #Saving Excel File
    time.sleep(5)
    navegador.find_element(By.ID,
        'formGeral:exportarExtratoExcel').click()

    #Closing Pou-up after download #Optimization Opportunity .By Method
    elemento = WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.XPATH,
        '/html/body/section/div/div/form/div[3]/div[8]/div[1]/a')))
    time.sleep(1) 
    elemento.click()

    #Finding the donwloaded file and saving in the target folder
    time.sleep(10)
    downloads_path = str(Path.home() / "Downloads")
    downloads_path = downloads_path + "\*"
    list_of_files = glob.glob(downloads_path) 
    latest_file = max(list_of_files, key=os.path.getctime)
    original = latest_file 
    shutil.move(original, target)

    #Closing Browser
    time.sleep(5)
    navegador.close()


