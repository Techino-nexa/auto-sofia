import os
import time
import undetected_chromedriver as uc
import pyautogui as pa
import pandas as pd

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from datetime import datetime


import threading

# uc = undetected_chromedriver object
usuario = "cla00005"
senha = "fiscal022026"

from datetime import date
import calendar

def criar_pastas():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    dir_path = os.path.join(base_dir, "DOWNLOADS PLACEHOLDER")
    dir_relatorios = os.path.join(base_dir, "RELATORIOS")

    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

    if not os.path.exists(dir_relatorios):
        os.makedirs(dir_relatorios)

def quebrar_por_mes(data_inicio, data_fim):

    data_inicio = datetime.strptime(data_inicio, "%d/%m/%Y").date()
    data_fim = datetime.strptime(data_fim, "%d/%m/%Y").date()

    resultado = []

    ano = data_inicio.year
    mes = data_inicio.month

    while (ano, mes) <= (data_fim.year, data_fim.month):

        primeiro_dia_mes = date(ano, mes, 1)
        ultimo_dia_mes = date(ano, mes, calendar.monthrange(ano, mes)[1])

        inicio = max(data_inicio, primeiro_dia_mes)
        fim = min(data_fim, ultimo_dia_mes)

        resultado.append([
            inicio.strftime("%d/%m/%Y"),
            fim.strftime("%d/%m/%Y")
        ])

        if mes == 12:
            mes = 1
            ano += 1
        else:
            mes += 1

    return resultado

def checar_nenhumregistro(driver):
    time.sleep(3)
    try:
        alert = driver.switch_to.alert
        text = f"({alert.text}) "
        alert.accept()
        return text
    except NoAlertPresentException:
        return False

def clicar(driver, xpath):
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
        driver.find_element('xpath', xpath).click()
    except:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
        elemento = driver.find_element('xpath', xpath)
        driver.execute_script("arguments[0].click();", elemento)

def escrever(driver, xpath, mensagem):
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", xpath)))
    try:
        driver.find_element("xpath", xpath).clear()
        driver.find_element("xpath", xpath).send_keys(mensagem)
    except:
        driver.find_element("xpath", xpath).send_keys(mensagem)

def entrar_no_sefaz(dir_download):
    uc_chrome_options = uc.ChromeOptions()
    dict_chrome_prefs = {
        "download.default_directory": dir_download,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    uc_chrome_options.add_experimental_option("prefs", dict_chrome_prefs)
    uc_chrome_options.add_argument("--force-device-scale-factor=0.6")


    driver = None
    driver = uc.Chrome(
        service=Service(ChromeDriverManager().install()), 
        options=uc_chrome_options,
        version_main=144
    )

    driver.get("https://sefaz.pb.gov.br/servirtual")

    

    iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')

    driver.switch_to.frame(iframeatf)
    try:
        escrever(driver, '//*[@id="form-cblogin-username"]/div/input', usuario)
    except:
        return("SEFAZ FORA DO AR 🔴")
    escrever(driver, '//*[@id="form-cblogin-password"]/div[1]/input', senha)
    clicar(driver, '//*[@id="form-cblogin-password"]/div[2]/input[2]')

    driver.switch_to.default_content()

    iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')

    driver.switch_to.frame(iframeatf)

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(('xpath', '/html/body/form/b/a')))
    element = driver.find_element('xpath', '/html/body/form/b/a')
    ActionChains(driver).key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()
        
    return(driver)

def entrar_no_sefaz_certificado(dir_download):
    uc_chrome_options = uc.ChromeOptions()
    dict_chrome_prefs = {
        "download.default_directory": dir_download,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    uc_chrome_options.add_experimental_option("prefs", dict_chrome_prefs)
    uc_chrome_options.add_argument("--force-device-scale-factor=0.6")


    driver = None
    driver = uc.Chrome(
        service=Service(ChromeDriverManager().install()), 
        options=uc_chrome_options,
        version_main=144
    )

    driver.get("https://sefaz.pb.gov.br/servirtual")

    

    iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')

    driver.switch_to.frame(iframeatf)
    try:
        escrever(driver, '//*[@id="form-cblogin-username"]/div/input', "CERTIFICADO")
    except:
        return("SEFAZ FORA DO AR 🔴")
    
    def func1():
        clicar(driver,'//*[@id="SERlogin"]/form/div[3]/div/a')

    def func2():
        time.sleep(5)
        pa.press("enter")
        print("UOU")
        print("função 2 rodando")
    
    threading.Thread(target=func1).start()
    threading.Thread(target=func2).start()

    driver.switch_to.default_content()

    iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')

    driver.switch_to.frame(iframeatf)

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(('xpath', '/html/body/form/b/a')))
    element = driver.find_element('xpath', '/html/body/form/b/a')
    ActionChains(driver).key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()
        
    return(driver)


def clicar_if(driver, xpath):
    try:
        driver.switch_to.default_content()
        driver.switch_to.frame("iframe")
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
        driver.find_element('xpath', xpath).click()
    except:
        time.sleep(1)
        driver.switch_to.default_content()
        driver.switch_to.frame("iframe")
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
        elemento = driver.find_element('xpath', xpath)
        driver.execute_script("arguments[0].click();", elemento)
        driver.switch_to.default_content()

def click_refresh(driver, x, id_download):
    try:
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))

        id_verificador = 3
        verificador = driver.find_element('xpath', f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[6]/a').text
        
        while verificador != id_download:
            id_verificador += 2
            verificador = driver.find_element('xpath',  f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[6]/a').text

        driver.find_element('xpath', f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[3]/a').click()
        return "CLICOU"
        
    except Exception as e:
        time.sleep(x)
        if x < 60:
            x += 1
        if x > 2:
            clicar(driver, f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[6]/a')
            try:
                try:
                    clicar(driver, "/html/body/form/div/table/tbody/tr[5]/td[4]/a")
                except:
                    clicar(driver, "/html/body/form/div/table/tbody/tr[3]/td[4]/a")
            except:
                driver.refresh()
                return click_refresh(driver, x, id_download)
            
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(('xpath', "/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]")))
            texto = driver.find_element('xpath', '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]').text
            
            if "Erro" in texto or "porém não há dados a serem exibidos" in texto:
                driver.switch_to.window(driver.window_handles[0])
                driver.refresh()
                driver.switch_to.window(driver.window_handles[1])
                driver.refresh()
                return texto
            else:
                driver.switch_to.window(driver.window_handles[0])
                driver.refresh()
                driver.switch_to.window(driver.window_handles[1])
                driver.refresh()
                return click_refresh(driver, x, id_download)
        else:
            driver.switch_to.window(driver.window_handles[0])
            driver.refresh()
            driver.switch_to.window(driver.window_handles[1])
            driver.refresh()
            return click_refresh(driver, x, id_download)

def esperar_download(pasta, timeout=120):
    tempo_inicial = time.time()

    while True:
        arquivos = os.listdir(pasta)

        # se tiver arquivo temporário do chrome ainda baixando
        if True in (".crdownload" in a.lower() for a in arquivos):
            pass
        else:
            if arquivos:
                return True
            
        if time.time() - tempo_inicial > timeout:
            return False
        time.sleep(1)



def download_nfe_emitente(dir_path,dict_infos):

    if dict_infos["certificado"]:
        driver = entrar_no_sefaz_certificado(dir_path)
    else:
        driver = entrar_no_sefaz(dir_path)

    if "SEFAZ FORA DO AR" in str(driver).upper():
        return ("SEFAZ FORA DO AR 🔴")
    
    driver.switch_to.default_content()
    driver.get("https://sefaz.pb.gov.br/servirtual/documentos-fiscais/nf-e/consulta-emitentes-destinatarios")
        
    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')
    
    clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/input[4]')
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', dict_infos["data_inicial"])
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', dict_infos["data_final"])
    clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]/td/table/tbody/tr[1]/td[2]/select/option[2]')

    iframeemitente = driver.find_element("xpath", '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]/td/table/tbody/tr[2]/td/iframe')
    driver.switch_to.frame(iframeemitente)
    escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', dict_infos["cnpj"])
    clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')
    
    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')

    try:
        clicar(driver, '//html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/select/option[3]')
    except:
        clicar(driver, '//html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/select/option[2]')
    
    clicar(driver, '//*[@id="btnConsulta"]')

    erro = checar_nenhumregistro(driver)
    if erro == False:
        erro = "NENHUMA TELA ENCONTRADA 👍"
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text

        verificador = click_refresh(driver, 1, id_download)
        if "erro" in verificador.lower() or "porém não há dados a serem exibidos" in verificador.lower():
            driver.switch_to.window(driver.window_handles[0])
            return str(verificador)
        else:
            clicar_if(driver, '/html/body/form/div/table/tbody/tr[5]/td[3]/a')
            clicar_if(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a')
        
        driver.switch_to.window(driver.window_handles[0])
    
    else:
        erro = "TELA DE NENHUM REGISTRO ENCONTRADA 🔴"
    
    time.sleep(5)
    esperar_download(pasta=dir_path)
    driver.quit()
    return erro

def download_nfe_destinatario(dir_path,dict_infos):

    if dict_infos["certificado"]:
        driver = entrar_no_sefaz_certificado(dir_path)
    else:
        driver = entrar_no_sefaz(dir_path)
    if "SEFAZ FORA DO AR" in str(driver).upper():
        return ("SEFAZ FORA DO AR 🔴")
    
    driver.switch_to.default_content()
    driver.get("https://sefaz.pb.gov.br/servirtual/documentos-fiscais/nf-e/consulta-emitentes-destinatarios")

    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')

    clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/input[4]')
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', dict_infos["data_inicial"])
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', dict_infos["data_final"])
    clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]/td/table/tbody/tr[1]/td[2]/select/option[2]')

    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')
    clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[10]/td/table/tbody/tr[1]/td[2]/select/option[2]')

    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')
    iframedestinario = driver.find_element('xpath',  '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[10]/td/table/tbody/tr[2]/td/iframe')
    driver.switch_to.frame(iframedestinario)
    escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', dict_infos['cnpj'])
    clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')
    
    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')

    try:
        clicar(driver, '//html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/select/option[3]')
    except:
        clicar(driver, '//html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/select/option[2]')

    clicar(driver, '//*[@id="btnConsulta"]')

    erro = checar_nenhumregistro(driver)

    if erro == False:
        erro = "NENHUMA TELA ENCONTRADA 👍"
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text

        verificador = click_refresh(driver, 1, id_download)
        if "erro" in verificador.lower() or "porém não há dados a serem exibidos" in verificador.lower():
            driver.switch_to.window(driver.window_handles[0])
            return str(verificador)
        else:
            clicar_if(driver, '/html/body/form/div/table/tbody/tr[5]/td[3]/a')
            clicar_if(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a')
        
        driver.switch_to.window(driver.window_handles[0])
    
    else:
        erro = "TELA DE NENHUM REGISTRO ENCONTRADA 🔴"
    
    time.sleep(5)
    esperar_download(pasta=dir_path)
    driver.quit()
    return erro

def download_nfce(dir_path,dict_infos):

    if dict_infos["certificado"]:
        driver = entrar_no_sefaz_certificado(dir_path)
    else:
        driver = entrar_no_sefaz(dir_path)

    if "SEFAZ FORA DO AR" in str(driver).upper():
        return ("SEFAZ FORA DO AR 🔴")
    

    driver.switch_to.default_content()
    driver.get("https://sefaz.pb.gov.br/servirtual/documentos-fiscais/nfc-e/consulta-por-emitente")
    
    driver.switch_to.frame('iframe')

    clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[11]/td[2]/input[3]')

    time.sleep(2)
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', dict_infos["data_inicial"])
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', dict_infos["data_final"])
    time.sleep(2)
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', dict_infos["data_inicial"])
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', dict_infos["data_final"])
    
    clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/table/tbody/tr[1]/td[2]/select/option[2]')

    iframeemitente = driver.find_element("xpath", '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/table/tbody/tr[2]/td/iframe')
    driver.switch_to.frame(iframeemitente)

    escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', dict_infos["cnpj"])
    clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')
    
    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')
    
    try:
        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[11]/td[2]/select[1]/option[3]')
    except:
        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[11]/td[2]/select[1]/option[2]')

    clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[11]/td[2]/button')

    erro = checar_nenhumregistro(driver)

    if erro == False:
        erro = "NENHUMA TELA ENCONTRADA 👍"
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text

        verificador = click_refresh(driver, 1, id_download)
        if "erro" in verificador.lower() or "porém não há dados a serem exibidos" in verificador.lower():
            driver.switch_to.window(driver.window_handles[0])
            return str(verificador)
        else:
            clicar_if(driver, '/html/body/form/div/table/tbody/tr[5]/td[3]/a')
            clicar_if(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a')
        
        driver.switch_to.window(driver.window_handles[0])
    
    else:
        erro = "TELA DE NENHUM REGISTRO ENCONTRADA 🔴"
    
    time.sleep(5)
    esperar_download(pasta=dir_path)
    driver.quit()
    return erro

def download_cte_destinatario(dir_path,dict_infos):
    if dict_infos["certificado"]:
        driver = entrar_no_sefaz_certificado(dir_path)
    else:
        driver = entrar_no_sefaz(dir_path)

    if "SEFAZ FORA DO AR" in str(driver).upper():
        return ("SEFAZ FORA DO AR 🔴")
    
    driver.switch_to.default_content()
    driver.get("https://sefaz.pb.gov.br/servirtual/documentos-fiscais/ct-e/gerar-xml-ct-e")
        
    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')
    
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[22]/td/input[5]')
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', dict_infos["data_inicial"])
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', dict_infos["data_final"])
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[6]/td[2]/select[1]/option[2]')
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[20]/td/table/tbody/tr[1]/td[2]/select/option[2]')  # DESTINATARIO
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[14]/td/table/tbody/tr[1]/td[2]/select/option[2]')  # TOMADOR
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[22]/td/select/option[2]')  # XML


    iframetomador = driver.find_element('xpath', '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[14]/td/table/tbody/tr[2]/td/iframe')
    driver.switch_to.frame(iframetomador)
    escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', dict_infos["cnpj"])
    clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')

    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')
    iframedestinatario = driver.find_element('xpath', '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[20]/td/table/tbody/tr[2]/td/iframe')
    driver.switch_to.frame(iframedestinatario)
    escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', dict_infos["cnpj"])
    clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')

    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')

    clicar(driver, '//*[@id="btnConsultar"]')
    erro = checar_nenhumregistro(driver)
    if erro == False:
        erro = "NENHUMA TELA ENCONTRADA 👍"
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text

        verificador = click_refresh(driver, 1, id_download)
        if "erro" in verificador.lower() or "porém não há dados a serem exibidos" in verificador.lower():
            driver.switch_to.window(driver.window_handles[0])
            return str(verificador)
        else:

            clicar_if(driver,"/html/body/form/div/table/tbody/tr[7]/td[3]/a/img")
            clicar_if(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a')
        driver.switch_to.window(driver.window_handles[0])
    
    else:
        erro = "TELA DE NENHUM REGISTRO ENCONTRADA 🔴"
    
    time.sleep(5)
    esperar_download(pasta=dir_path)
    driver.quit()
    return erro

def download_cte_emitente(dir_path,dict_infos):

    if dict_infos["certificado"]:
        driver = entrar_no_sefaz_certificado(dir_path)
    else:
        driver = entrar_no_sefaz(dir_path)

    if "SEFAZ FORA DO AR" in str(driver).upper():
        return ("SEFAZ FORA DO AR 🔴")
    
    driver.switch_to.default_content()
    driver.get("https://sefaz.pb.gov.br/servirtual/documentos-fiscais/ct-e/gerar-xml-ct-e")
        
    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[22]/td/input[5]')

    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', dict_infos["data_inicial"])
    escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', dict_infos["data_final"])
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[6]/td[2]/select[1]/option[2]')
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[20]/td/table/tbody/tr[1]/td[2]/select/option[2]')  # DESTINATARIO
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[11]/td/table/tbody/tr[1]/td[2]/select/option[2]')  # EMITENTE
    clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[22]/td/select/option[2]')  # XML

    iframedestinatario = driver.find_element('xpath', '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[20]/td/table/tbody/tr[2]/td/iframe')
    driver.switch_to.frame(iframedestinatario)
    escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', dict_infos["cnpj"])
    clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')

    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')
    iframeremitente = driver.find_element('xpath', '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[11]/td/table/tbody/tr[2]/td/iframe')
    driver.switch_to.frame(iframeremitente)
    escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', dict_infos["cnpj"])
    clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')

    driver.switch_to.default_content()
    driver.switch_to.frame('iframe')

    clicar(driver, '//*[@id="btnConsultar"]')

    erro = checar_nenhumregistro(driver)
    if erro == False:
        erro = "NENHUMA TELA ENCONTRADA 👍"
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text

        verificador = click_refresh(driver, 1, id_download)
        if "erro" in verificador.lower() or "porém não há dados a serem exibidos" in verificador.lower():
            driver.switch_to.window(driver.window_handles[0])
            return str(verificador)
        else:
            clicar_if(driver,"/html/body/form/div/table/tbody/tr[7]/td[3]/a/img")
            clicar_if(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a')
        
        driver.switch_to.window(driver.window_handles[0])
    
    else:
        erro = "TELA DE NENHUM REGISTRO ENCONTRADA 🔴"
    
    time.sleep(5)
    esperar_download(pasta=dir_path)
    driver.quit()
    return erro



if __name__ == "__main__":
    criar_pastas()
    
    datas = (quebrar_por_mes("01/02/2020", "31/12/2025"))

    dict_relatorio = {
        "PERIODO":[],
        "EMPRESA":[],
        "CNPJ":[],
        "STATUS":[],
    }

    base_dir = os.path.dirname(os.path.abspath(__file__))
    dir_path = os.path.join(base_dir, "DOWNLOADS PLACEHOLDER")


    base_dir = os.path.dirname(os.path.abspath(__file__))
    dir_relatorios = os.path.join(base_dir, "RELATORIOS")

    for data in datas:

        dict_infos = {
            "cnpj":"09158643000132",
            "data_inicial":data[0].replace("/",""),
            "data_final":data[1].replace("/",""),
            "certificado":True,
        }
        resultado = download_nfe_emitente(dir_path,dict_infos)

        dict_relatorio["PERIODO"].append(data[0])
        dict_relatorio["EMPRESA"].append("LOJAO")
        dict_relatorio["CNPJ"].append(dict_infos["cnpj"])
        dict_relatorio["STATUS"].append(resultado)


    
        df = pd.DataFrame(dict_relatorio)
        df.to_excel(fr"{dir_relatorios}/relatorio.xlsx", index=False)
    




