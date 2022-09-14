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
import datetime
from datetime import timedelta
from config import *

config = CatalogConfig()
config.read()

def open_table_mkt():

    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_argument('headless')
    prefs = {"download.default_directory" : config['PATH']['PB']}
    chromeOptions.add_experimental_option("prefs",prefs)
    driver = webdriver.Chrome(options=chromeOptions, service=Service(ChromeDriverManager().install()))

    driver.get("https://guestcenter.opentable.com/login")
    time.sleep(10)
    inputElement = driver.find_element(by=By.XPATH, value='//*[@id="email"]')
    inputElement.send_keys(config['OPEN']['EMAIL'])
    inputElement.send_keys(Keys.ENTER)
    time.sleep(5)
    inputElement = driver.find_element(by=By.XPATH, value='//*[@id="okta-signin-password"]')
    inputElement.send_keys(config['OPEN']['PASSWORD'])
    inputElement.send_keys(Keys.ENTER)

    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/nav/div/div/ul/li[2]/button/div[1]/div[2]'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="chrome"]/div/div[1]/div/button/span'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="chrome"]/div/div[1]/div/div[2]/div[2]/div/section[1]/ul/li[8]/a/span'))).click()
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/main/section/header/div/aside/div[1]/button/div/div/span'))).click()
    time.sleep(15)

    data = datetime.date.today() - timedelta(1)
    nome_mes = data.strftime("%B")
    nome_dia = data.strftime("%A")
    dia = data.strftime("%d")
    mes = data.strftime("%m")
    ano = data.year

    data_valor_investido = f"Choose {nome_dia}, {nome_mes} {dia}, {ano} as your check-in date. It’s available."
    data_value= f"//*[@aria-label='{data_valor_investido}']"
    print(data_valor_investido)

    source = driver.find_element(by=By.XPATH, value=data_value)
    action = ActionChains(driver)
    action.double_click(source).perform()

    time.sleep(15)

    driver.find_element(By.CLASS_NAME, "Button__buttonContent___2ZgNMUuizmm4SieYLp96xV").click()

    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/main/section/div/div/article[2]/header/button/span/span'))).click()

    time.sleep(20)

    directory = Path(config['PATH']['PB'])
    lista = os.listdir(directory)
    print(lista)
    files = directory.glob('*.csv')
    arquivo_mais_recente = max(files, key=getmtime)
    print(arquivo_mais_recente)

    df = pd.read_csv(arquivo_mais_recente, sep = ',')

    lista = []
    for i in df['Gasto total']:
        lista.append(i)
    lista

    total_invested = lista[-1].split()[-1]
    dt = f'{ano}-{mes}-{dia}'

    df_periodo = pd.read_excel(config['PATH']['OPEN'], sheet_name = 'Periodo')
    df_periodo = df_periodo.fillna(0)
    df_periodo = df_periodo[['Período', 'Data', 'Valor Investido']]
    df_periodo['Valor Investido'][df_periodo['Data'] == dt] = float(total_invested.replace(',', '.'))
    os.remove(arquivo_mais_recente)
    return df_periodo
