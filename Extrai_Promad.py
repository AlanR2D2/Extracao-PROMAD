from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import traceback
import pandas as pd
from dotenv import load_dotenv
from monitoring import log
import send_email
from tenacity import retry, stop_after_attempt, wait_fixed
import tempfile

load_dotenv()

email = os.getenv("EMAIL")
senha = os.getenv("SENHA")

log.info("Inicializando ChromeDriver")
path_chromedriver = ChromeDriverManager().install()
path_chromedriver = path_chromedriver.replace('THIRD_PARTY_NOTICES.chromedriver', 'chromedriver.exe')
path_chromedriver = path_chromedriver.replace('LICENSE.chromedriver', 'chromedriver.exe')
servico = Service(path_chromedriver)

options = webdriver.ChromeOptions()
temp_dir = tempfile.mkdtemp()
options.add_argument(f"--user-data-dir={temp_dir}")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--headless=new")

log.info("Criando navegador com opções headless")
nav = webdriver.Chrome(service=servico, options=options)
wait = WebDriverWait(nav, 10)

log.info('Navegador criado com sucesso')

def scroll_to_element(driver, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element)
        time.sleep(1)
    except Exception as e:
        log.error(f"Erro ao realizar scroll até o elemento: {e}")

@retry(stop=stop_after_attempt(3), wait=wait_fixed(5))
def get_data():
    try:
        log.info("Acessando página de login")
        nav.get("https://www.integra.adv.br/login-integra.asp")

        login_label = wait.until(EC.presence_of_element_located((By.NAME, 'txtUsuario')))
        login_label.send_keys(email)

        senha_label = wait.until(EC.presence_of_element_located((By.NAME, 'txtSenha')))
        senha_label.send_keys(senha)

        sj_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form_login"]/div[5]/div/div/label[1]/input')))
        sj_button.click()

        button_access = wait.until(EC.element_to_be_clickable((By.ID, 'btn-acessar-conta')))
        button_access.click()

        log.info('Login realizado com sucesso')

        time.sleep(4)
        nav.get('https://www.integra.adv.br/moderno/modulo/50/default.asp')
        log.info("Página de relatórios acessada")

        time.sleep(3)

        botao_fase = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div[2]/div[21]/div[1]/button')))
        nav.execute_script("arguments[0].click();", botao_fase)
        log.info("Filtro de 'Fase' aberto")

        todos_opt = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div[2]/div[21]/div[1]/div/div/ul/li[1]/a/span[2]')))
        scroll_to_element(nav, todos_opt)
        todos_opt.click()
        log.info("Filtro 'Todos' aplicado")

        pesquisar_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnPesquisar"]')))
        scroll_to_element(nav, pesquisar_button)
        pesquisar_button.click()
        log.info("Botão 'Pesquisar' clicado")

        excel_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frmRelatorio"]/div[3]/div[1]/div/ul/li[5]/div/span')))
        scroll_to_element(nav, excel_button)
        excel_button.click()
        log.info("Relatório Excel selecionado")

        modelo_slct = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/button')))
        scroll_to_element(nav, modelo_slct)
        modelo_slct.click()

        modelo_label = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/div[1]/div/div/input')))
        modelo_label.send_keys('CONTROLE ATUALIZADO')

        contr_atualizado = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/div[1]/ul/li[2]/label')))
        contr_atualizado.click()
        log.info("Modelo de relatório 'CONTROLE ATUALIZADO' selecionado")

        time.sleep(10)

        gerar_excel = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnIlimitado"]')))
        scroll_to_element(nav, gerar_excel)
        gerar_excel.click()
        log.info("Download do relatório iniciado")

        time.sleep(50)
        log.info('get_data finalizada com sucesso')

    except Exception as e:
        msg = f'Ocorreu erro na def get_data na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. ERROR: {e}'
        log.error(msg)
        raise Exception("Algo em get_data. Iniciando outra tentativa")
