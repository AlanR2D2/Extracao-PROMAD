from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
import pandas as pd
import os
from datetime import datetime
import time
import shutil
import re
import traceback
import pandas as pd
import os
from dotenv import load_dotenv
from monitoring import Logger
import Trata_arquivos
import send_email
from tenacity import retry, stop_after_attempt, wait_fixed

load_dotenv()

email = os.getenv("EMAIL")
senha = os.getenv("SENHA")

log = Logger('PROMAD', silencer=False).get_logger()

path_chromedriver = ChromeDriverManager().install()
path_chromedriver = path_chromedriver.replace('THIRD_PARTY_NOTICES.chromedriver', 'chromedriver.exe')
path_chromedriver = path_chromedriver.replace('LICENSE.chromedriver', 'chromedriver.exe')
servico = Service (path_chromedriver)
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")
#options.add_argument("--window-size=1920,1080")
nav = webdriver.Chrome(service=servico, options=options)
wait = WebDriverWait(nav, 10)  # Espera máxima de 10 segundos

log.info('Navegador criado com sucesso')

def scroll_to_element(driver, element):
    """
    Realiza o scroll até o elemento especificado.

    :param driver: Instância do WebDriver.
    :param by: Estratégia de localização (ex: By.ID, By.XPATH).
    :param value: Valor correspondente à estratégia de localização.
    """
    try:
        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element)
        time.sleep(1)  # Aguarda para garantir que o scroll foi concluído
    except Exception as e:
        print(f"Erro ao realizar scroll até o elemento: {e}")

@retry(stop=stop_after_attempt(3), wait=wait_fixed(5))
def get_data():
    """
    Realiza login e navegação automatizada no sistema Promad via Selenium, seleciona filtros e exporta o relatório
    "CONTROLE ATUALIZADO" em formato Excel.

    Funcionalidades principais:
    ---------------------------
    - Acessa a página de login do sistema Promad.
    - Realiza login com as credenciais fornecidas via variáveis globais.
    - Navega até a área de relatórios.
    - Seleciona o filtro de "Fase" e marca a opção "Todos".
    - Clica em "Pesquisar" para carregar os dados.
    - Seleciona o modelo de relatório "CONTROLE ATUALIZADO".
    - Gera e inicia o download do relatório em Excel.

    Observações:
    ------------
    - Usa `WebDriverWait` para aguardar os elementos estarem presentes ou clicáveis.
    - Utiliza `execute_script` como fallback para cliques que falham por sobreposição de elementos.
    - Há um `time.sleep(50)` após o clique no botão "Gerar Excel" para garantir que o download seja concluído.
    - Em caso de erro, o log é registrado, uma mensagem de e-mail é enviada e a execução é encerrada.
    """

    try:
        nav.get("https://www.integra.adv.br/login-integra.asp")

        # Aguarda o campo de login estar presente e envia o e-mail
        login_label = wait.until(EC.presence_of_element_located((By.NAME, 'txtUsuario')))
        login_label.send_keys(email)

        # Aguarda o campo de senha estar presente e envia a senha
        senha_label = wait.until(EC.presence_of_element_located((By.NAME, 'txtSenha')))
        senha_label.send_keys(senha)

        # Aguarda o botão de seleção estar clicável e clica
        sj_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form_login"]/div[5]/div/div/label[1]/input')))
        sj_button.click()

        # Aguarda o botão de acesso estar clicável e clica
        button_access = wait.until(EC.element_to_be_clickable((By.ID, 'btn-acessar-conta')))
        button_access.click()

        log.info('Login realizado')

        #Acessa página de relatórios
        time.sleep(4)
        nav.get('https://www.integra.adv.br/moderno/modulo/50/default.asp')

        time.sleep(3)  # Ajuste conforme necessário

        # Clica em Fase

        botao_fase = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div[2]/div[21]/div[1]/button')))
        nav.execute_script("arguments[0].click();", botao_fase)

        #Clica em Todos
        todos_opt = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div[2]/div[21]/div[1]/div/div/ul/li[1]/a/span[2]')))
        scroll_to_element(nav, todos_opt)
        todos_opt.click()

        #Clica em Pesquisar
        pesquisar_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnPesquisar"]')))
        scroll_to_element(nav, pesquisar_button)
        pesquisar_button.click()

        #Clica em Relatório Excel
        excel_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frmRelatorio"]/div[3]/div[1]/div/ul/li[5]/div/span')))
        scroll_to_element(nav, excel_button)
        excel_button.click()

        #Clica em Relatório Excel
        modelo_slct = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/button')))
        scroll_to_element(nav, modelo_slct)
        modelo_slct.click()

        modelo_label = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/div[1]/div/div/input')))
        modelo_label.send_keys('CONTROLE ATUALIZADO')

        contr_atualizado = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/div[1]/ul/li[2]/label')))
        contr_atualizado.click()

        time.sleep(10)

        gerar_excel = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnIlimitado"]')))
        scroll_to_element(nav, gerar_excel)
        gerar_excel.click()

        time.sleep(50)

        log.info('get_data finalizada')
    except Exception as e:
        msg = f'Ocorreu erro na def get_data na linha {traceback.extract_tb(e2.__traceback__)[0].lineno}'
        print(msg)
        log.error (msg)
        send_email.send_email_error(msg)
        exit()


