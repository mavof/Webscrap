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
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException

#Manage Files
from pathlib import Path 
import glob
import os
import shutil

#Time Sleep control
import time

#Send notifications
from plyer import notification

#Interacting with excel files
import pandas as pd

servico = Service(ChromeDriverManager().install())

#Run the chrome in 2. Plan
from selenium.webdriver.chrome.options import Options

#Functions to make the browser wait for elements
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#Function to take care of downloading file
def enable_download_headless(browser,download_dir):
    browser.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
    browser.execute("send_command", params)

# instantiate a chrome options object so you can set the size and headless preference
# some of these chrome options might be uncessary but I just used a boilerplate
# change the <path_to_download_default_directory> to whatever your default download folder is located

downloads_path = str(Path.home() / "Downloads")
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920x1080")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--verbose')
chrome_options.add_experimental_option("prefs", {
        "download.default_directory": downloads_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False
})
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--disable-software-rasterizer')

# initialize driver object and change the <path_to_chrome_driver> depending on your directory where your chromedriver should be
navegador = webdriver.Chrome(options=chrome_options, service=servico)

# change the <path_to_place_downloaded_file> to your directory where you would like to place the downloaded file
download_dir = downloads_path #"<path_to_place_downloaded_file>"

# function to handle setting up headless download
enable_download_headless(navegador, download_dir)

#Read the information to download the files
tabela = pd.read_excel("Reference.xlsx")

#Run the routine for each line in the information source file
for linha in tabela.index:
    Bank = str(tabela.loc[linha,"Bank"])
    Cond = str(tabela.loc[linha,"Cond"])
    Ag = str(tabela.loc[linha,"AG"])
    CC = str(tabela.loc[linha,"CC"])
    Login = str(tabela.loc[linha,"Login"])
    Password = str(tabela.loc[linha,"Password"])
    Target = str(tabela.loc[linha,"Target"])
    
    WebDriverWaitTime = 30 #Time reference to commonize the WebDriverWait function
    i=0
    while i<1:
            #Open the website
        navegador.get("https://www.santander.com.br/empresas")
        time.sleep(4)

        #Access 
        navegador.find_element(By.NAME,'txtAgencia').clear()
        time.sleep(1)
        navegador.find_element(By.NAME,'txtAgencia').send_keys(Ag)
        
        navegador.find_element(By.NAME,'txtConta').clear()
        time.sleep(1)
        navegador.find_element(By.NAME,'txtConta').send_keys(CC, Keys.ENTER)
        
        #Access
        time.sleep(15)
        navegador.find_element(By.ID,'formGeral:usuario').clear()
        time.sleep(1)
        navegador.find_element(By.ID,'formGeral:usuario').send_keys(Login)
        
        navegador.find_element(By.ID,'formGeral:senhaAcesso').clear()
        time.sleep(1)
        navegador.find_element(By.ID,'formGeral:senhaAcesso').send_keys(Password, Keys.ENTER)

        #Closing the introduction Pou-up
        try:
            if WebDriverWait(navegador, WebDriverWaitTime).until(EC.presence_of_element_located((By.ID,
                                                                                                 'closeMessage'))).is_displayed:
                    elemento = WebDriverWait(navegador, WebDriverWaitTime).until(EC.presence_of_element_located((By.ID,
                                                                                                                 'closeMessage')))
                    time.sleep(1) # Avoiding errors
                    elemento.click()
                    notification.notify(
                        title = 'Notificação',
                        message = 'A execução irá aguardar 5 minutos para dar continuidade',
                        app_icon = None,
                        timeout = 10,
                    )
                    navegador.close()
                    time.sleep(320)
                    navegador = webdriver.Chrome(options=chrome_options, service=servico)
                    download_dir = downloads_path
                    enable_download_headless(navegador, download_dir)
                    continue
        except: TimeoutException
        i=1
    notification.notify(
        title = 'Notificação',
        message = 'Login '+ Cond + ' - ' + Bank +' realizado',
        app_icon = None,
        timeout = 10,
    )
    try:
        elemento = WebDriverWait(navegador, WebDriverWaitTime).until(EC.presence_of_element_located((By.ID,
            'formGeral:btnCloseModalLightBox')))
        time.sleep(1) # Avoiding errors
        elemento.click()
    except: NoSuchElementException

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

    #find the botton of the page
    time.sleep(1)
    navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    #Saving .PDF File
    time.sleep(5)
    navegador.find_element(By.ID,
        'formGeral:salvarPDF').click()

    #Closing Pou-up after download #Optimization Opportunity .By Method
    elemento = WebDriverWait(navegador, WebDriverWaitTime).until(EC.presence_of_element_located((By.XPATH,
        '/html/body/section/div/div/form/div[3]/div[8]/div[1]/a')))
    time.sleep(1) # Avoiding errors
    elemento.click()

    try:
        #Finding the donwloaded file and saving in the target folder
        time.sleep(10)
        downloads_path = str(Path.home() / "Downloads")
        downloads_path = downloads_path + "\*"
        list_of_files = glob.glob(downloads_path) 
        latest_file = max(list_of_files, key=os.path.getctime)
        original = latest_file 
        shutil.move(original, Target)
    except:
        notification.notify(
        title = 'Notificação',
        message = 'O arquivo não foi salvo na pasta de destino. Verifique a pasta de downloads',
        app_icon = None,
        timeout = 10,
    )

    #Saving Excel File
    time.sleep(5)
    navegador.find_element(By.ID,
        'formGeral:exportarExtratoExcel').click()

    #Closing Pou-up after download #Optimization Opportunity .By Method
    elemento = WebDriverWait(navegador, WebDriverWaitTime).until(EC.presence_of_element_located((By.XPATH,
        '/html/body/section/div/div/form/div[3]/div[8]/div[1]/a')))
    time.sleep(1) # Avoiding errors
    elemento.click()
    
    try:
        #Finding the donwloaded file and saving in the target folder
        time.sleep(10)
        downloads_path = str(Path.home() / "Downloads")
        downloads_path = downloads_path + "\*"
        list_of_files = glob.glob(downloads_path) 
        latest_file = max(list_of_files, key=os.path.getctime)
        original = latest_file 
        shutil.move(original, target)
    except:
        notification.notify(
        title = 'Notificação',
        message = 'O arquivo não foi salvo na pasta de destino. Verifique a pasta de downloads',
        app_icon = None,
        timeout = 10,
    )

    #Closing Browser
    time.sleep(5)
    navegador.close()
    
    notification.notify(
        title = 'Notificação',
        message = 'Execução ' + Cond + ' - ' + Bank + ' concluída',
        app_icon = None,
        timeout = 10,
    )


