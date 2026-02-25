# sistemas/auto_sofia.py
import os
import time
import sys
import django
import undetected_chromedriver as uc
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from datetime import datetime

from pathlib import Path

def baixar_cte(per_inicial, per_final, cnpj, senha, cte_tomador=False, cte_emitente=False, download_dir=None, headless=False, usuario="cla0000"):
    
    erros = ""
    
    # Define diretório de download
    if download_dir is None:
        download_dir = getattr(settings, 'SOFIA_DOWNLOAD_DIR', os.path.join(settings.BASE_DIR, 'downloads', 'sofia'))
    
    os.makedirs(download_dir, exist_ok=True)
    
    print(f"[{datetime.now()}] CTE - Iniciando download para CNPJ: {cnpj}")
    print(f"[{datetime.now()}] CTE - Período: {per_inicial} a {per_final}")
    print(f"[{datetime.now()}] CTE - Tomador: {'Sim' if cte_tomador else 'Não'} | Emitente: {'Sim' if cte_emitente else 'Não'}")
    
    if not cte_tomador and not cte_emitente:
        print(f"[{datetime.now()}] CTE - Nenhum tipo de CTe selecionado. Pulando...")
        return ""
    
    # --------------------------------------------------------------
    # FUNÇÕES AUXILIARES
    # --------------------------------------------------------------
    
    def click_refresh(driver, x, id_download):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))

            id_verificador = 3
            verificador = driver.find_element('xpath', f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[6]/a').text
            
            while verificador != id_download:
                id_verificador += 2
                verificador = driver.find_element('xpath', f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[6]/a').text

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
                    if x > 10:
                        return "ERRO"
                    else:
                        driver.refresh()
                        return click_refresh(driver, x, id_download)
                
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable(('xpath', "/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]")))
                texto = driver.find_element('xpath', '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]').text
                print(f"[{datetime.now()}] CTE - Mensagem SEFAZ: {texto}")
                
                if "Erro" in texto or "porém não há dados a serem exibidos" in texto:
                    driver.switch_to.window(driver.window_handles[0])
                    driver.refresh()
                    driver.switch_to.window(driver.window_handles[1])
                    driver.refresh()
                    return "ERRO"
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

    def clicar(driver, xpath):
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
            driver.find_element('xpath', xpath).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
            elemento = driver.find_element('xpath', xpath)
            driver.execute_script("arguments[0].click();", elemento)

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

    def escrever(driver, xpath, mensagem):
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", xpath)))
        try:
            driver.find_element("xpath", xpath).clear()
            driver.find_element("xpath", xpath).send_keys(mensagem)
        except:
            driver.find_element("xpath", xpath).send_keys(mensagem)

    def checar_nenhumregistro(driver):
        time.sleep(3)
        try:
            alert = driver.switch_to.alert
            text = f"({alert.text}) "
            alert.accept()
            return text
        except NoAlertPresentException:
            return "0"

    def download_tomador(driver): #download 1
        print(f"[{datetime.now()}] CTE - Processando download Tomador/Destinatário")
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')
        
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text
        
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')
        
        checar = click_refresh(driver, 1, id_download)
        
        if checar != "ERRO":
            clicar_if(driver, '/html/body/form/div/table/tbody/tr[7]/td[3]/a')
            clicar_if(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a')
        
        driver.switch_to.window(driver.window_handles[0])

    def download_emitente(driver): #download 2
        print(f"[{datetime.now()}] CTE - Processando download Emitente/Destinatário")
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text

        print(f"[{datetime.now()}] CTE - ID Download Emitente: {id_download}")
        
        checar = click_refresh(driver, 1, id_download)
        
        if checar != "ERRO":
            clicar_if(driver, '/html/body/form/div/table/tbody/tr[7]/td[3]/a')
            clicar_if(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a')
        
        driver.switch_to.window(driver.window_handles[0])

    # --------------------------------------------------------------
    # CONFIGURAÇÃO DO CHROME
    # --------------------------------------------------------------
    
    chrome_options = uc.ChromeOptions()
    chrome_prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", chrome_prefs)
    chrome_options.add_argument("--force-device-scale-factor=0.6")
    
    if headless:
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = None
    
    try:
        driver = uc.Chrome(
            service=Service(ChromeDriverManager().install()), 
            options=chrome_options,
            version_main=144
        )
        
        # --------------------------------------------------------------
        # LOGIN NO SEFAZ
        # --------------------------------------------------------------
        
        print(f"[{datetime.now()}] CTE - Acessando SEFAZ-PB")
        driver.get("https://sefaz.pb.gov.br/servirtual")
        
        iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')
        driver.switch_to.frame(iframeatf)
        escrever(driver, '//*[@id="form-cblogin-username"]/div/input', usuario)
        escrever(driver, '//*[@id="form-cblogin-password"]/div[1]/input', senha)
        
        clicar(driver, '//*[@id="form-cblogin-password"]/div[2]/input[2]')
        print(f"[{datetime.now()}] CTE - Login realizado")

        driver.switch_to.default_content()
        iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')
        driver.switch_to.frame(iframeatf)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(('xpath', '/html/body/form/b/a')))
        element = driver.find_element('xpath', '/html/body/form/b/a')
        ActionChains(driver).key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()
        
        # Padroniza CNPJ para 14 dígitos
        if len(cnpj) == 13:
            cnpj = f"0{cnpj}"
        
        # --------------------------------------------------------------
        # ACESSA PÁGINA DE CTE
        # --------------------------------------------------------------
        
        driver.switch_to.default_content()
        driver.get("https://sefaz.pb.gov.br/servirtual/documentos-fiscais/ct-e/gerar-xml-ct-e")
        print(f"[{datetime.now()}] CTE - Acessou página de consulta")
        
        # --------------------------------------------------------------
        # DOWNLOAD 1: TOMADOR/DESTINATÁRIO
        # --------------------------------------------------------------
        
        if cte_tomador:
            print(f"[{datetime.now()}] CTE - Iniciando busca Tomador/Destinatário")
            
            driver.switch_to.window(driver.window_handles[1])
            driver.switch_to.window(driver.window_handles[0])
            driver.switch_to.frame('iframe')

            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[22]/td/input[5]')
            escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', per_inicial)
            escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', per_final)
            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[6]/td[2]/select[1]/option[2]')
            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[20]/td/table/tbody/tr[1]/td[2]/select/option[2]')  # DESTINATARIO
            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[14]/td/table/tbody/tr[1]/td[2]/select/option[2]')  # TOMADOR
            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[22]/td/select/option[2]')  # XML

            iframetomador = driver.find_element('xpath', '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[14]/td/table/tbody/tr[2]/td/iframe')
            driver.switch_to.frame(iframetomador)
            escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', cnpj)
            clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')

            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            iframedestinatario = driver.find_element('xpath', '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[20]/td/table/tbody/tr[2]/td/iframe')
            driver.switch_to.frame(iframedestinatario)
            escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', cnpj)
            clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')

            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            clicar(driver, '//*[@id="btnConsultar"]')

            erro = checar_nenhumregistro(driver)
            if erro != "0":
                erros += f"CTe Tomador: {erro}"
                print(f"[{datetime.now()}] CTE - Erro Tomador: {erro}")
            
            if str(erro) == "0":
                download_tomador(driver)
                print(f"[{datetime.now()}] CTE - Download Tomador concluído")
            else:
                erros += "CTe Tomador - Nenhum registro encontrado\n"
        
        # --------------------------------------------------------------
        # DOWNLOAD 2: EMITENTE/DESTINATÁRIO
        # --------------------------------------------------------------
        
        if cte_emitente:
            print(f"[{datetime.now()}] CTE - Iniciando busca Emitente/Destinatário")
            
            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[22]/td/input[5]')

            escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', per_inicial)
            escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', per_final)
            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[6]/td[2]/select[1]/option[2]')
            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[20]/td/table/tbody/tr[1]/td[2]/select/option[2]')  # DESTINATARIO
            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[11]/td/table/tbody/tr[1]/td[2]/select/option[2]')  # EMITENTE
            clicar(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[22]/td/select/option[2]')  # XML

            iframedestinatario = driver.find_element('xpath', '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[20]/td/table/tbody/tr[2]/td/iframe')
            driver.switch_to.frame(iframedestinatario)
            escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', cnpj)
            clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')

            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            iframeremitente = driver.find_element('xpath', '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[11]/td/table/tbody/tr[2]/td/iframe')
            driver.switch_to.frame(iframeremitente)
            escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', cnpj)
            clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')

            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            clicar(driver, '//*[@id="btnConsultar"]')
            
            erro = checar_nenhumregistro(driver)
            if erro != "0":
                erros += f"CTe Emitente: {erro}"
                print(f"[{datetime.now()}] CTE - Erro Emitente: {erro}")
            
            if str(erro) == "0":
                download_emitente(driver)
                print(f"[{datetime.now()}] CTE - Download Emitente concluído")
            else:
                erros += "CTe Emitente - Nenhum registro encontrado\n"

            driver.switch_to.default_content()
            driver.refresh()

        time.sleep(5)
        print(f"[{datetime.now()}] CTE - Processo finalizado")
        
    except Exception as e:
        erro_msg = f"Erro crítico em baixar_cte: {str(e)}"
        print(f"[{datetime.now()}] CTE ERRO - {erro_msg}")
        erros += erro_msg + "\n"
        
    finally:
        if driver:
            try:
                time.sleep(10)
                driver.quit()
                print(f"[{datetime.now()}] CTE - Driver fechado")
            except:
                pass

    return erros

def baixar_nfe(per_inicial, per_final, cnpj, senha, nfe_emitente=False, nfe_destinatario=False, download_dir=None, headless=False, usuario="cla0000"):
    
    print(nfe_destinatario, nfe_emitente)
    erros = ""

    # Define diretório de download
    if download_dir is None:
        download_dir = getattr(settings, 'SOFIA_DOWNLOAD_DIR', os.path.join(settings.BASE_DIR, 'downloads', 'sofia'))
    
    os.makedirs(download_dir, exist_ok=True)
    
    print(f"[{datetime.now()}] NFE - Iniciando download para CNPJ: {cnpj}")
    print(f"[{datetime.now()}] NFE - Período: {per_inicial} a {per_final}")
    print(f"[{datetime.now()}] NFE - Emitente: {'Sim' if nfe_emitente else 'Não'} | Destinatário: {'Sim' if nfe_destinatario else 'Não'}")
    
    # Se nenhum foi marcado, não faz nada
    if not nfe_emitente and not nfe_destinatario:
        print(f"[{datetime.now()}] NFE - Nenhum tipo de NFe selecionado. Pulando...")
        return ""
    # --------------------------------------------------------------
    # FUNÇÕES AUXILIARES
    # --------------------------------------------------------------
    
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
                print(f"[{datetime.now()}] NFE - Mensagem SEFAZ: {texto}")
                
                if "Erro" in texto or "porém não há dados a serem exibidos" in texto:
                    driver.switch_to.window(driver.window_handles[0])
                    driver.refresh()
                    driver.switch_to.window(driver.window_handles[1])
                    driver.refresh()
                    return "ERRO"
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

    def clicar(driver, xpath):
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
            driver.find_element('xpath', xpath).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
            elemento = driver.find_element('xpath', xpath)
            driver.execute_script("arguments[0].click();", elemento)

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

    def escrever(driver, xpath, mensagem):
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", xpath)))
        try:
            driver.find_element("xpath", xpath).clear()
            driver.find_element("xpath", xpath).send_keys(mensagem)
        except:
            driver.find_element("xpath", xpath).send_keys(mensagem)

    def checar_nenhumregistro(driver):
        time.sleep(3)
        try:
            alert = driver.switch_to.alert
            text = f"({alert.text}) "
            alert.accept()
            return text
        except NoAlertPresentException:
            return "0"

    def download_emitente(driver):
        print(f"[{datetime.now()}] NFE - Processando download Emitente")
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text

        print(f"[{datetime.now()}] NFE - ID Download Emitente: {id_download}")

        verificador = click_refresh(driver, 1, id_download)
        
        if verificador == "ERRO":
            return "ERRO DOWNLOAD NFE EMITENTE\n"
        else:
            clicar_if(driver, '/html/body/form/div/table/tbody/tr[5]/td[3]/a')
            clicar_if(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a')
            print(f"[{datetime.now()}] NFE - Download Emitente concluído")
        
        driver.switch_to.window(driver.window_handles[0])
        return ""

    def download_destinatario(driver):
        print(f"[{datetime.now()}] NFE - Processando download Destinatário")
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        WebDriverWait(driver, 10).until( EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text

        print(f"[{datetime.now()}] NFE - ID Download Destinatário: {id_download}")

        verificador = click_refresh(driver, 1, id_download)
        
        if verificador == "ERRO":
            return "ERRO DOWNLOAD NFE DESTINATÁRIO\n"
        else:
            clicar_if(driver, '/html/body/form/div/table/tbody/tr[5]/td[3]/a')
            clicar_if(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a')
            print(f"[{datetime.now()}] NFE - Download Destinatário concluído")
        
        driver.switch_to.window(driver.window_handles[0])
        return ""

    def preencher_formulario_base(driver):
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')
        
        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/input[4]')
        escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', per_inicial)
        escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', per_final)
        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]/td/table/tbody/tr[1]/td[2]/select/option[2]')

    # --------------------------------------------------------------
    # CONFIGURAÇÃO DO CHROME
    # --------------------------------------------------------------
    
    chrome_options = uc.ChromeOptions()
    chrome_prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", chrome_prefs)
    chrome_options.add_argument("--force-device-scale-factor=0.6")
    
    if headless:
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = None
    
    try:
        driver = uc.Chrome(
            service=Service(ChromeDriverManager().install()), 
            options=chrome_options,
            version_main=144
        )
        
        # --------------------------------------------------------------
        # LOGIN NO SEFAZ
        # --------------------------------------------------------------
        
        print(f"[{datetime.now()}] NFE - Acessando SEFAZ-PB")
        driver.get("https://sefaz.pb.gov.br/servirtual")
        
        iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')
        driver.switch_to.frame(iframeatf)
        escrever(driver, '//*[@id="form-cblogin-username"]/div/input', usuario)
        escrever(driver, '//*[@id="form-cblogin-password"]/div[1]/input', senha)
        
        clicar(driver, '//*[@id="form-cblogin-password"]/div[2]/input[2]')
        print(f"[{datetime.now()}] NFE - Login realizado")

        driver.switch_to.default_content()
        iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')
        driver.switch_to.frame(iframeatf)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(('xpath', '/html/body/form/b/a')))
        element = driver.find_element('xpath', '/html/body/form/b/a')
        ActionChains(driver).key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()
        
        # Padroniza CNPJ para 14 dígitos
        if len(cnpj) == 13:
            cnpj = f"0{cnpj}"
        
        # --------------------------------------------------------------
        # ACESSA PÁGINA DE NFE
        # --------------------------------------------------------------
        
        driver.switch_to.default_content()
        driver.get("https://sefaz.pb.gov.br/servirtual/documentos-fiscais/nf-e/consulta-emitentes-destinatarios")
        print(f"[{datetime.now()}] NFE - Acessou página de consulta")
        
        # --------------------------------------------------------------
        # DOWNLOAD 1: EMITENTE
        # --------------------------------------------------------------
        
        if nfe_emitente == True:
            print(f"[{datetime.now()}] NFE - Iniciando busca Emitente")
            
            preencher_formulario_base(driver)

            iframeemitente = driver.find_element("xpath", '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]/td/table/tbody/tr[2]/td/iframe')
            driver.switch_to.frame(iframeemitente)
            escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', cnpj)
            clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')
            
            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')

            try:
                clicar(driver, '//html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/select/option[3]')
            except:
                clicar(driver, '//html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/select/option[2]')
            
            clicar(driver, '//*[@id="btnConsulta"]')

            erro = checar_nenhumregistro(driver)
            if erro != "0":
                erros += f"NFe Emitente: {erro}"
                print(f"[{datetime.now()}] NFE - Erro Emitente: {erro}")
            
            if str(erro) == "0":
                erro_download = download_emitente(driver)
                if erro_download:
                    erros += erro_download
            else:
                erros += "NFe Emitente - Nenhum registro encontrado\n"
        
        # --------------------------------------------------------------
        # DOWNLOAD 2: DESTINATÁRIO
        # --------------------------------------------------------------
        
        if nfe_destinatario == True:
            print(f"[{datetime.now()}] NFE - Iniciando busca Destinatário")
            
            preencher_formulario_base(driver)

            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[10]/td/table/tbody/tr[1]/td[2]/select/option[2]')

            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            iframedestinario = driver.find_element('xpath',  '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[10]/td/table/tbody/tr[2]/td/iframe')
            driver.switch_to.frame(iframedestinario)
            escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', cnpj)
            clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')
            
            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')

            try:
                clicar(driver, '//html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/select/option[3]')
            except:
                clicar(driver, '//html/body/table/tbody/tr[2]/td/form/table/tbody/tr[12]/td/select/option[2]')

            clicar(driver, '//*[@id="btnConsulta"]')
            
            erro = checar_nenhumregistro(driver)
            if erro != "0":
                erros += f"NFe Destinatário: {erro}"
                print(f"[{datetime.now()}] NFE - Erro Destinatário: {erro}")
            
            if str(erro) == "0":
                erro_download = download_destinatario(driver)
                if erro_download:
                    erros += erro_download
            else:
                erros += "NFe Destinatário - Nenhum registro encontrado\n"

        time.sleep(6)
        print(f"[{datetime.now()}] NFE - Aguardando downloads...")
        print(f"[{datetime.now()}] NFE - Processo finalizado")
        
    except Exception as e:
        erro_msg = f"Erro crítico em baixar_nfe: {str(e)}"
        print(f"[{datetime.now()}] NFE ERRO - {erro_msg}")
        erros += erro_msg + "\n"
        
    finally:
        if driver:
            try:
                time.sleep(10)
                driver.quit()
                print(f"[{datetime.now()}] NFE - Driver fechado")
            except:
                pass

    return erros

def baixar_nfce(per_inicial, per_final, cnpj, senha, download_dir=None, headless=False, usuario="cla0000"):

    erros = ""
    
    # Define diretório de download
    if download_dir is None:
        download_dir = getattr(settings, 'SOFIA_DOWNLOAD_DIR', os.path.join(settings.BASE_DIR, 'downloads', 'sofia'))
    
    os.makedirs(download_dir, exist_ok=True)
    
    print(f"[{datetime.now()}] NFCE - Iniciando download para CNPJ: {cnpj}")
    print(f"[{datetime.now()}] NFCE - Período: {per_inicial} a {per_final}")
    
    # --------------------------------------------------------------
    # FUNÇÕES AUXILIARES
    # --------------------------------------------------------------
    
    def click_refresh(driver, x, id_download):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))

            id_verificador = 3
            verificador = driver.find_element('xpath', f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[6]/a').text
            
            while verificador != id_download:
                id_verificador += 2
                verificador = driver.find_element('xpath', f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[6]/a').text

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
                
                if "Erro" in texto or "porém não há dados a serem exibidos" in texto or "agendado" in texto.lower():
                    print(f"[{datetime.now()}] NFCE ERRO SEFAZ - Mensagem: {texto}")
                    driver.switch_to.window(driver.window_handles[0])
                    driver.refresh()
                    driver.switch_to.window(driver.window_handles[1])
                    driver.refresh()
                    return "ERRO"
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

    def clicar(driver, xpath):
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
            driver.find_element('xpath', xpath).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
            elemento = driver.find_element('xpath', xpath)
            driver.execute_script("arguments[0].click();", elemento)
            print(f"[{datetime.now()}] NFCE - Tentativa alternativa de clique: {xpath}")

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
            print(f"[{datetime.now()}] NFCE - Tentativa alternativa de clique: {xpath}")

    def escrever(driver, xpath, mensagem):
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", xpath)))
        try:
            driver.find_element("xpath", xpath).clear()
            driver.find_element("xpath", xpath).send_keys(mensagem)
        except:
            driver.find_element("xpath", xpath).send_keys(mensagem)

    def checar_nenhumregistro(driver):
        time.sleep(3)
        try:
            alert = driver.switch_to.alert
            text = f"({alert.text}) "
            alert.accept()
            return text
        except NoAlertPresentException:
            return "0"

    def download_nfce(driver):
        print(f"[{datetime.now()}] NFCE - Processando download")
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()

        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))
        id_download = driver.find_element("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text
        print(f"[{datetime.now()}] NFCE - ID Download: {id_download}")

        verificador = click_refresh(driver, 1, id_download)
        
        if verificador == "ERRO":
            print(f"[{datetime.now()}] NFCE - Falha no download - ID: {id_download}")
            return "ERRO DOWNLOAD NFCE\n"
        else:
            print(f"[{datetime.now()}] NFCE - Download iniciado - ID: {id_download}")
            clicar_if(driver, "/html/body/form/div/table/tbody/tr[5]/td[3]/a")
            
            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')

            clicar(driver, "/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/a")
            print(f"[{datetime.now()}] NFCE - Download concluído")

        driver.switch_to.window(driver.window_handles[0])
        driver.switch_to.default_content()
        return ""

    # --------------------------------------------------------------
    # CONFIGURAÇÃO DO CHROME
    # --------------------------------------------------------------
    
    chrome_options = uc.ChromeOptions()
    chrome_prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", chrome_prefs)
    chrome_options.add_argument("--force-device-scale-factor=0.6")
    
    if headless:
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = None
    
    try:
        driver = uc.Chrome(
            service=Service(ChromeDriverManager().install()), 
            options=chrome_options,
            version_main=144
        )
        
        # --------------------------------------------------------------
        # LOGIN NO SEFAZ
        # --------------------------------------------------------------
        
        print(f"[{datetime.now()}] NFCE - Acessando SEFAZ-PB")
        driver.get("https://sefaz.pb.gov.br/servirtual")
        
        iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')
        driver.switch_to.frame(iframeatf)
        escrever(driver, '//*[@id="form-cblogin-username"]/div/input', usuario)
        escrever(driver, '//*[@id="form-cblogin-password"]/div[1]/input', senha)
        
        clicar(driver, '//*[@id="form-cblogin-password"]/div[2]/input[2]')
        print(f"[{datetime.now()}] NFCE - Login realizado")

        driver.switch_to.default_content()
        iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')
        driver.switch_to.frame(iframeatf)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(('xpath', '/html/body/form/b/a')))
        element = driver.find_element('xpath', '/html/body/form/b/a')
        ActionChains(driver).key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()
        print(f"[{datetime.now()}] NFCE - Caixa de mensagens aberta")
        
        # Padroniza CNPJ para 14 dígitos
        if len(cnpj) == 13:
            cnpj = f"0{cnpj}"
        
        # --------------------------------------------------------------
        # ACESSA PÁGINA DE NFCE
        # --------------------------------------------------------------
        
        driver.switch_to.default_content()
        driver.get("https://sefaz.pb.gov.br/servirtual/documentos-fiscais/nfc-e/consulta-por-emitente")
        print(f"[{datetime.now()}] NFCE - Acessou página de consulta")
        
        # --------------------------------------------------------------
        # PREENCHE FORMULÁRIO
        # --------------------------------------------------------------
        
        driver.switch_to.frame('iframe')

        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[11]/td[2]/input[3]')

        time.sleep(2)
        escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', per_inicial)
        escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[1]', per_inicial)
        time.sleep(2)
        escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', per_final)
        escrever(driver, '/html/body/table[1]/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/input[2]', per_final)
        print(f"[{datetime.now()}] NFCE - Período definido: {per_inicial} a {per_final}")
        
        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/table/tbody/tr[1]/td[2]/select/option[2]')

        iframeemitente = driver.find_element("xpath", '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[8]/td/table/tbody/tr[2]/td/iframe')
        driver.switch_to.frame(iframeemitente)

        escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', cnpj)
        print(f"[{datetime.now()}] NFCE - CNPJ definido: {cnpj}")
        clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')
        
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')
        
        try:
            clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[11]/td[2]/select[1]/option[3]')
        except:
            clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[11]/td[2]/select[1]/option[2]')

        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[11]/td[2]/button')
        print(f"[{datetime.now()}] NFCE - Busca iniciada")
        
        # --------------------------------------------------------------
        # VERIFICA RESULTADO E FAZ DOWNLOAD
        # --------------------------------------------------------------
        
        erro = checar_nenhumregistro(driver)
        if erro != "0":
            print(f"[{datetime.now()}] NFCE - Erro: {erro}")
            erros += erro
        else:
            print(f"[{datetime.now()}] NFCE - Registros encontrados")
        
        if str(erro) == "0":
            erro_download = download_nfce(driver)
            if erro_download:
                erros += erro_download
            time.sleep(5)
        else:
            erros += "NFCe - Nenhum registro encontrado\n"

        time.sleep(6)
        print(f"[{datetime.now()}] NFCE - Aguardando downloads...")
        print(f"[datetime.now()] NFCE - Processo finalizado")
        
    except Exception as e:
        erro_msg = f"Erro crítico em baixar_nfce: {str(e)}"
        print(f"[{datetime.now()}] NFCE ERRO - {erro_msg}")
        erros += erro_msg + "\n"
        
    finally:
        if driver:
            try:
                time.sleep(10)
                driver.quit()
                print(f"[{datetime.now()}] NFCE - Driver fechado")
            except:
                pass

    print(f"[{datetime.now()}] NFCE - Erros acumulados: {erros if erros else 'Nenhum'}")
    return erros

def baixar_faturamento(per_inicial, per_final, cnpj, senha, download_dir=None, headless=False, usuario="cla0000"):
    
    erros = ""
    
    # Define diretório de download
    if download_dir is None: download_dir = getattr(settings, 'SOFIA_DOWNLOAD_DIR', os.path.join(settings.BASE_DIR, 'downloads', 'sofia'))
    
    os.makedirs(download_dir, exist_ok=True)
    
    print(f"[{datetime.now()}] FATURAMENTO - Iniciando download para CNPJ: {cnpj}")
    print(f"[{datetime.now()}] FATURAMENTO - Período: {per_inicial} a {per_final}")
    
    # --------------------------------------------------------------
    # FUNÇÕES AUXILIARES
    # --------------------------------------------------------------
    
    def click_refresh(driver, x, id_download):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame('iframe')
            WebDriverWait(driver, 1).until(EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a')))

            id_verificador = 3
            verificador = driver.find_element('xpath', f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[6]/a').text
            
            while verificador != id_download:
                id_verificador += 2
                verificador = driver.find_element('xpath', f'/html/body/form/div/table/tbody/tr[{id_verificador}]/td[6]/a').text

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
                texto = driver.find_element('xpath', 
                    '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[7]').text
                print(f"[{datetime.now()}] FATURAMENTO - Mensagem SEFAZ: {texto}")
                
                if "Erro" in texto or "porém não há dados a serem exibidos" in texto:
                    driver.switch_to.window(driver.window_handles[0])
                    driver.refresh()
                    driver.switch_to.window(driver.window_handles[1])
                    driver.refresh()
                    return "ERRO"
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

    def clicar(driver, xpath):
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
            driver.find_element('xpath', xpath).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(("xpath", xpath)))
            elemento = driver.find_element('xpath', xpath)
            driver.execute_script("arguments[0].click();", elemento)

    def clicar_if(driver, xpath):
        """Clica em elemento dentro de iframe"""
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

    def clicar_link(driver, texto_link):
        """Clica em um link pelo texto"""
        elemento_link = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.LINK_TEXT, texto_link))
        )
        driver.execute_script("arguments[0].click();", elemento_link)

    def escrever(driver, xpath, mensagem):
        """Escreve em um campo"""
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(("xpath", xpath)))
        try:
            driver.find_element("xpath", xpath).clear()
            driver.find_element("xpath", xpath).send_keys(mensagem)
        except:
            driver.find_element("xpath", xpath).send_keys(mensagem)

    def checar_nenhumregistro(driver):
        """Verifica se apareceu alerta de 'nenhum registro'"""
        time.sleep(3)
        try:
            alert = driver.switch_to.alert
            text = f"({alert.text}) "
            alert.accept()
            return text
        except NoAlertPresentException:
            return "0"

    def processar_download(driver):
        """Processa o download do relatório de faturamento"""
        print(f"[{datetime.now()}] FATURAMENTO - Processando download")
        
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(("xpath", '/html/body/form/div/table/tbody/tr[3]/td[6]/a'))
        )
        id_download = driver.find_element("xpath", 
            '/html/body/form/div/table/tbody/tr[3]/td[6]/a').text

        print(f"[{datetime.now()}] FATURAMENTO - ID Download: {id_download}")

        verificador = click_refresh(driver, 1, id_download)
        
        if verificador == "ERRO":
            print(f"[{datetime.now()}] FATURAMENTO - Erro no processamento")
            return "ERRO"
        else:
            clicar(driver, "/html/body/form/div/table/tbody/tr[5]/td[3]/a")
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame('iframe')
                clicar_link(driver, "Relatório de Consulta do Faturamento Diário NFC-e")
                print(f"[{datetime.now()}] FATURAMENTO - Download concluído")
                return "OK"
            except Exception as e:
                print(f"[{datetime.now()}] FATURAMENTO - Tentando novamente...")
                # Recursão: tenta novamente
                return processar_download(driver)

    # --------------------------------------------------------------
    # CONFIGURAÇÃO DO CHROME
    # --------------------------------------------------------------
    
    chrome_options = uc.ChromeOptions()
    chrome_prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", chrome_prefs)
    chrome_options.add_argument("--force-device-scale-factor=0.6")
    
    if headless:
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
    
    driver = None
    
    try:
        driver = uc.Chrome(
            service=Service(ChromeDriverManager().install()), 
            options=chrome_options,
            version_main=144
        )
        
        # --------------------------------------------------------------
        # LOGIN NO SEFAZ
        # --------------------------------------------------------------
        
        print(f"[{datetime.now()}] FATURAMENTO - Acessando SEFAZ-PB")
        driver.get("https://sefaz.pb.gov.br/servirtual")
        
        iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')
        driver.switch_to.frame(iframeatf)
        escrever(driver, '//*[@id="form-cblogin-username"]/div/input', usuario)
        escrever(driver, '//*[@id="form-cblogin-password"]/div[1]/input', senha)
        
        clicar(driver, '//*[@id="form-cblogin-password"]/div[2]/input[2]')
        print(f"[{datetime.now()}] FATURAMENTO - Login realizado")

        driver.switch_to.default_content()
        iframeatf = driver.find_element("xpath", '//*[@id="atf-login"]/iframe')
        driver.switch_to.frame(iframeatf)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '/html/body/form/b/a'))
        )
        element = driver.find_element('xpath', '/html/body/form/b/a')
        ActionChains(driver).key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()
        print(f"[{datetime.now()}] FATURAMENTO - Caixa de mensagens aberta")
        
        # Padroniza CNPJ para 14 dígitos
        if len(cnpj) == 13:
            cnpj = f"0{cnpj}"
        
        # --------------------------------------------------------------
        # NAVEGA PELO MENU ATÉ FATURAMENTO
        # --------------------------------------------------------------
        
        driver.switch_to.default_content()
        
        print(f"[{datetime.now()}] FATURAMENTO - Navegando pelo menu")
        clicar(driver, '//*[@id="menu-servirtual"]')
        clicar(driver, '//*[@id="off-menu_215"]/div[2]/div[1]/dl/dt[10]')
        clicar(driver, '//*[@id="off-menu_215"]/div[2]/div[1]/dl/dd[10]/div/dl/dt[5]/div/div')
        clicar(driver, '//*[@id="off-menu_215"]/div[2]/div[1]/dl/dd[10]/div/dl/dd[5]/div/dl/dt[8]/div/div/a')
        print(f"[{datetime.now()}] FATURAMENTO - Acessou página de consulta")
        
        # --------------------------------------------------------------
        # PREENCHE FORMULÁRIO
        # --------------------------------------------------------------
        
        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table[1]/tbody[2]/tr/td/table/tbody/tr[2]/td[2]/select/option[2]')
        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table[2]/tbody[2]/tr[2]/td[2]/input[2]')
        escrever(driver, '/html/body/table/tbody/tr[2]/td/form/table[2]/tbody[2]/tr[1]/td[2]/input[1]', per_inicial)
        escrever(driver, '/html/body/table/tbody/tr[2]/td/form/table[2]/tbody[2]/tr[1]/td[2]/input[2]', per_final)
        print(f"[{datetime.now()}] FATURAMENTO - Período definido: {per_inicial} a {per_final}")

        iframecriterio = driver.find_element('xpath', 
            '/html/body/table/tbody/tr[2]/td/form/table[1]/tbody[2]/tr/td/table/tbody/tr[3]/td/iframe')
        driver.switch_to.frame(iframecriterio)

        escrever(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[2]/input', cnpj)
        print(f"[{datetime.now()}] FATURAMENTO - CNPJ definido: {cnpj}")
        clicar(driver, '//*[@id="Layer1"]/table/tbody/tr/td/form/table/tbody/tr[1]/td[3]/input')

        driver.switch_to.default_content()
        driver.switch_to.frame("iframe")

        clicar(driver, '/html/body/table/tbody/tr[2]/td/form/table[3]/tbody/tr/td/input[3]')
        print(f"[{datetime.now()}] FATURAMENTO - Busca iniciada")
        
        # --------------------------------------------------------------
        # PROCESSA DOWNLOAD
        # --------------------------------------------------------------
        
        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()

        driver.switch_to.default_content()
        driver.switch_to.frame('iframe')

        resultado = processar_download(driver)
        
        if resultado == "ERRO":
            erros += "Faturamento - Erro ao processar download\n"
            print(f"[{datetime.now()}] FATURAMENTO - Erro ao processar download")
        else:
            print(f"[{datetime.now()}] FATURAMENTO - Download concluído com sucesso")

        driver.switch_to.window(driver.window_handles[0])
        
        try:
            driver.switch_to.default_content()
            time.sleep(5)
        except:
            erros += "FATURAMENTO - Erro ao retornar à janela principal\n"
        
        time.sleep(6)
        print(f"[{datetime.now()}] FATURAMENTO - Aguardando downloads...")
        print(f"[{datetime.now()}] FATURAMENTO - Processo finalizado")
        
    except Exception as e:
        erro_msg = f"Erro crítico em baixar_faturamento: {str(e)}"
        print(f"[{datetime.now()}] FATURAMENTO ERRO - {erro_msg}")
        erros += erro_msg + "\n"
        
    finally:
        if driver:
            try:
                time.sleep(10)
                driver.quit()
                print(f"[{datetime.now()}] FATURAMENTO - Driver fechado")
            except:
                pass

    return erros

def executar_sofia_completo(dict_infos, per_inicial, per_final, senha="fiscal012026", usuario="cla00005", headless=False, download_dir=r"E:\Projetos\4 - SOFIA\DOWNLOADS"):

    # ========== NFe ==========
    if dict_infos["nfe1"] or dict_infos["nfe2"]:
        print(f"\nBaixando NFe...")
        erros_nfe = baixar_nfe(
            per_inicial, per_final, str(dict_infos["cnpj"]),
            senha, dict_infos["nfe1"], dict_infos["nfe2"],
            download_dir=download_dir, headless=headless, usuario=usuario
        )

    # ========== NFCe ==========
    if dict_infos["nfce"]:
        print(f"\nBaixando NFCe...")
        erros_nfce = baixar_nfce(
            per_inicial=per_inicial, per_final=per_final,
            cnpj=dict_infos["cnpj"], senha=senha,
            download_dir=download_dir, headless=headless, usuario=usuario
        )

    # ========== CTe ==========
    if dict_infos["cte2"] or dict_infos["cte1"]:
        print(f"\nBaixando CTe...")
        erros_cte = baixar_cte(
            per_inicial=per_inicial, per_final=per_final,
            cnpj=dict_infos["cnpj"], senha=senha,
            cte_tomador=dict_infos["cte2"], cte_emitente=dict_infos["cte1"],
            download_dir=download_dir, headless=headless, usuario=usuario
        )

    # ========== Faturamento ==========
    if dict_infos["faturamento"]:
        print(f"\nBaixando Faturamento...")
        erros_fat = baixar_faturamento(
            per_inicial=per_inicial, per_final=per_final,
            cnpj=dict_infos["cnpj"], senha=senha,
            download_dir=download_dir, headless=headless, usuario=usuario
        )
        
    # ========== REGISTRA ARQUIVOS BAIXADOS ==========
    documentos_criados = 0


    
        
    

