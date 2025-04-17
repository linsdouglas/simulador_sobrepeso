#Automação SAP
import time
import math
import  datetime
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinter import ttk
import customtkinter as ctk
from PIL import Image
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
import pyautogui as pt
import yagmail
from openpyxl import load_workbook
import comtypes.client
import glob
import shutil
import subprocess
import pandas as pd
import os
import sys

base_sap="base_sap"
download_dir = os.path.join(os.environ['USERPROFILE'], 'Downloads', base_sap)
if not os.path.exists(download_dir):
    os.makedirs(download_dir)
url = "https://s4.mdiasbranco.com.br:44380/sap/bc/gui/sap/its/webgui#"
driver = None
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920x1080")
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "directory_upgrade": True
}
chrome_options.add_experimental_option("prefs", prefs)


def login_sap():
    try:
        username_field = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='text' and @name='username']"))
        )
        print("Campo de usuário encontrado, efetuando printin...")
        username_field.clear()
        username_field.send_keys("xql80316")
        password_field = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='password' and @name='password']"))
        )
        password_field.clear()
        password_field.send_keys("8583Doug#")
        password_field.send_keys(Keys.ENTER)
        print("printin efetuado com sucesso.")
        time.sleep(5)
    except Exception as e:
        print("Login não requerido após o refresh ou elemento não encontrado")

def interacoes_sap(driver,actions):
    driver.set_window_size(1920, 2000)
    try:
        WebDriverWait(driver, 300).until(lambda d: d.execute_script("return document.readyState") == "complete")
        time.sleep(2)
        elemento_pesq = WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, 'TXT_TITULO')))
        safe_click(driver,(By.ID, 'ToolbarOkCode'),nome_elemento="barra de pesquisa principal do sap")
        time.sleep(1)
        elemento_pesq.clear()
        elemento_pesq.send_keys("/nZPM003")
        time.sleep(1)
        elemento_centro_trabalho = WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, 'M0:46:::4:34')))
        safe_click(driver,(By.ID, 'M0:46:::4:34'),nome_elemento="centro de trabalho")
        time.sleep(1)
        elemento_centro_trabalho.send_keys("M431")
        hoje = datetime.date.today()
        data_inicial = hoje - datetime.timedelta(days=30)
        data_inicial_input = driver.find_element(By.ID, "M0:46:::2:34")
        data_inicial_input.send_keys(data_inicial.strftime("%d.%m.%Y"))
        data_final_input = driver.find_element(By.ID, "M0:46:::2:59")
        data_final_input.send_keys(hoje.strftime("%d.%m.%Y"))
        executar_button = driver.find_element(By.ID, "M0:50::btn[8]")
        executar_button.click()
        time.sleep(3)
        WebDriverWait(driver, 300).until(lambda d: d.execute_script("return document.readyState") == "complete")
        elemento_pesq = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'grid#C114#4,4#if-r')))
        actions.key_down(Keys.SHIFT).send_keys(Keys.F9).key_up(Keys.SHIFT).perform()       
        elemento_relatorio = WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, 'M1:46:::2:3')))
        safe_click(driver,(By.ID, 'M1:50::btn[0]'),nome_elemento="exportar o relatório")
        WebDriverWait(driver, 300).until(lambda d: d.execute_script("return document.readyState") == "complete")
        elemento_confirmar = WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, 'webguiPopupWindow10000-header-title-txt')))
        safe_click(driver,(By.ID, 'UpDownDialogChoose'),nome_elemento="confirmar exportação")
        wait = WebDriverWait(driver, 20)  # tempo máximo de espera em segundos
        msg_element = wait.until(
            EC.presence_of_element_located((By.ID, "wnd[0]/sbar_msg-txt"))
        )
        mensagem = msg_element.text
        print("Mensagem recebida:", mensagem)
        if "Download" in mensagem and "EXPORT.XLSX" in mensagem:
            print("Download confirmado. Seguindo para o próximo passo.")
        else:
            print("Mensagem diferente da esperada:", mensagem)

        driver.find_element(By.ID, "u25056").click()
        safe_click(driver,(By.ID,'u25056'),nome_elemento="configuração pra abrir o relatório de downloads")
        wait = WebDriverWait(driver, 10)
        browser_item = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(text(), 'Browser de arquivo SAP GUI for HTML')]")
        ))
        safe_click(driver,(By.XPATH, "//div[contains(text(), 'Browser de arquivo SAP GUI for HTML')]"),nome_elemento="relatórios de download")
        WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, 'popupDialogExportBtn')))
        safe_click(driver,(By.ID,'popupDialogExportBtn'),nome_elemento="botão final para exportar")

def safe_click(driver, by_locator,nome_elemento="Elemento", timeout=10):
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable(by_locator)
        )
        element.click()
        print(f"Clique padrão realizado com sucesso no elemento:{nome_elemento}")
    except (ElementClickInterceptedException, TimeoutException) as e:
        print(f"Clique padrão falhou: {repr(e)}. Tentando com JavaScript...")
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located(by_locator)
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", element)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", element)
            print("Clique forçado via JavaScript realizado com sucesso.")
        except Exception as js_e:
            print(f"Erro ao clicar com JavaScript: {repr(js_e)}")
    except Exception as e:
        print(f"Erro inesperado ao tentar clicar: {repr(e)}")