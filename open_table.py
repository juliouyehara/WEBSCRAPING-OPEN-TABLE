import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import os
import pandas as pd
import xlsxwriter
from pathlib import Path
from os.path import getmtime
from config import *

config = CatalogConfig()
config.read()

def open_table_relatorio():

    options = Options()
    options.headless = True

    driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))
    driver.get("https://guestcenter.opentable.com/login")
    time.sleep(5)
    inputElement = driver.find_element(by=By.XPATH, value='//*[@id="email"]')
    inputElement.send_keys(config['OPEN']['EMAIL'])
    inputElement.send_keys(Keys.ENTER)
    time.sleep(3)
    inputElement = driver.find_element(by=By.XPATH, value='//*[@id="okta-signin-password"]')
    inputElement.send_keys(config['OPEN']['PASSWORD'])
    inputElement.send_keys(Keys.ENTER)
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/nav/div/div/ul/li[2]/button/div[1]/div[2]'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="chrome"]/div/div[1]/div/button/span'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="chrome"]/div/div[1]/div/div[2]/div[2]/div/section[1]/ul/li[6]/a/span'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="mainnav"]/nav/ul/li[3]/ul/li[3]'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="mainnav"]/nav/ul/li[3]/ul/li[3]'))).click()

    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[1]/div/div/button/div/div/span'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[1]/div/div/div/div/div/div[1]/div/button[2]/span'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[1]/div/div/div/div/div/div[2]/div[1]/div[1]/button[6]'))).click()

    time.sleep(25)

    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/button/div/span'))).click()

    ## Limpar filtros

    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/div[2]/button/span'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/div[2]/button/span'))).click()

    ## Selecionar filtros

    ###### Data visita
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[1]/ul/li[1]/button/span'))).click()
    ###### Data criação
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[1]/ul/li[2]/button/span'))).click()
    ###### Nome cliente
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[2]/ul/li[1]/button/span'))).click()
    ###### Tamanho
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[3]/ul/li[1]/button/span'))).click()
    ###### Status
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[3]/ul/li[2]/button/span'))).click()
    ###### Fonte
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[3]/ul/li[7]/button/span'))).click()
    ###### Parceiro
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[3]/ul/li[9]/button/span'))).click()
    ###### Tipo de descoberta
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[3]/ul/li[10]/button/span'))).click()
    ###### Fonte RestRef
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[3]/ul/li[12]/button/span'))).click()
    ###### Nome Campanha RestRef
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[3]/ul/li[13]/button/span'))).click()
    ###### Subtotal da experiencia
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[4]/ul/li[3]/button/span'))).click()
    ###### Receita total da experiencia
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[4]/ul/li[5]/button/span'))).click()
    ###### Receita total da experiencia 2
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[4]/ul/li[9]/button/span'))).click()
    ###### Receita no PDV
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[5]/ul/li[1]/button/span'))).click()
    ###### Gorjeta no PDV
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[5]/ul/li[2]/button/span'))).click()
    ###### IDS no PDV
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[5]/ul/li[3]/button/span'))).click()
    ###### Data transferencia stripe
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[6]/ul/li[6]/button/span'))).click()
    ###### Receita total
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[6]/ul/li[10]/button/span'))).click()
    ###### Gorjeta total
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/ul/li[6]/ul/li[11]/button/span'))).click()

    ###### Aplicar
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/div/section/div/div[3]/div[1]/div/div[6]/div/div/div/div/div[3]/button/span'))).click()

    time.sleep(15)

    driver.find_element(By.CLASS_NAME, "ActionButton-module__actionLabel___R9IgAemvYUxee7vhDddG").click()
    time.sleep(10)
    driver.quit()

    time.sleep(180)

    chromeOptions = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": config['PATH']['PATH']}
    chromeOptions.add_experimental_option("prefs", prefs)
    chromeOptions.add_argument('headless')

    driver = webdriver.Chrome(options=chromeOptions, service=Service(ChromeDriverManager().install()))
    driver.get("https://webmail.gestaoboomer.com/")
    window_before = driver.window_handles[0]
    time.sleep(5)
    inputElement = driver.find_element(by=By.XPATH, value='//*[@id="user"]')
    inputElement.send_keys(config['WEBMAIL']['EMAIL'])

    time.sleep(3)
    inputElement = driver.find_element(by=By.XPATH, value='//*[@id="pass"]')
    inputElement.send_keys(config['WEBMAIL']['PASSWORD'])
    inputElement.send_keys(Keys.ENTER)

    time.sleep(10)

    lista = []
    for row in driver.find_elements(by=By.XPATH, value='.//tr'):
        r = row.get_attribute('id')
        lista.append(str(r))
    id_xpath = lista[2]

    source = driver.find_element(by=By.XPATH, value=f'//*[@id="{id_xpath}"]/td[4]/span/span')

    action = ActionChains(driver)
    action.double_click(source).perform()

    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="message-htmlpart1"]/div/center/table/tbody/tr[2]/td/center/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr/td/div/table[2]/tbody/tr/td/a/span'))).click()

    time.sleep(15)

    driver.switch_to.window(driver.window_handles[-1])

    inputElement = driver.find_element(by=By.XPATH, value='//*[@id="email"]')
    inputElement.send_keys(config['OPEN']['EMAIL'])
    inputElement.send_keys(Keys.ENTER)
    time.sleep(3)
    inputElement = driver.find_element(by=By.XPATH, value='//*[@id="okta-signin-password"]')
    inputElement.send_keys(config['OPEN']['PASSWORD'])
    inputElement.send_keys(Keys.ENTER)

    time.sleep(30)

    directory = Path(config['PATH']['PB'])
    lista = os.listdir(directory)
    files = directory.glob('*.csv')
    arquivo_mais_recente = max(files, key=getmtime)

    csv = arquivo_mais_recente
    df = pd.read_csv(csv, sep=',')
    os.remove(csv)
    driver.quit()

    return df

