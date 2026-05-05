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
from tenacity import retry, stop_after_attempt, wait_fixed
import tempfile
import shutil

load_dotenv()


class ScrapperError(Exception):
    """Exceção customizada para erros do scrapper — não mata o processo."""
    pass


# --- Estado do navegador (inicializado sob demanda) ---
_nav = None
_wait = None
_temp_dir = None
_download_dir = None


def get_download_dir():
    """Retorna o diretório dedicado de downloads do Chrome."""
    return _download_dir


def init_browser():
    """Cria o navegador Chrome. Reutiliza se já existir."""
    global _nav, _wait, _temp_dir, _download_dir

    if _nav is not None:
        return _nav, _wait

    log.info("Inicializando ChromeDriver")
    path_chromedriver = ChromeDriverManager().install()
    path_chromedriver = path_chromedriver.replace('THIRD_PARTY_NOTICES.chromedriver', 'chromedriver.exe')
    path_chromedriver = path_chromedriver.replace('LICENSE.chromedriver', 'chromedriver.exe')
    servico = Service(path_chromedriver)

    _temp_dir = tempfile.mkdtemp()
    _download_dir = tempfile.mkdtemp(prefix="promad_downloads_")

    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={_temp_dir}")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--headless=new")

    prefs = {"download.default_directory": _download_dir}
    options.add_experimental_option("prefs", prefs)

    log.info("Criando navegador com opções headless")
    _nav = webdriver.Chrome(service=servico, options=options)
    _wait = WebDriverWait(_nav, 10)
    log.info('Navegador criado com sucesso')

    return _nav, _wait


def quit_browser():
    """Fecha o navegador e limpa diretórios temporários."""
    global _nav, _wait, _temp_dir, _download_dir

    if _nav is not None:
        try:
            _nav.quit()
            log.info("Navegador fechado com sucesso")
        except Exception as e:
            log.warning(f"Erro ao fechar navegador: {e}")
        _nav = None
        _wait = None

    for d in (_temp_dir, _download_dir):
        if d and os.path.exists(d):
            try:
                shutil.rmtree(d, ignore_errors=True)
            except Exception:
                pass
    _temp_dir = None
    _download_dir = None


def scroll_to_element(driver, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", element)
        time.sleep(1)
    except Exception as e:
        log.error(f"Erro ao realizar scroll até o elemento: {e}")


@retry(stop=stop_after_attempt(3), wait=wait_fixed(5))
def get_data(status: str, email: str = None, senha: str = None, filtro_nome: str = 'CONTROLE ATUALIZADO'):
    '''
    :param status recebe 'Ativo' ou 'Inativo'
    :param email login do sistema
    :param senha senha do sistema
    :param filtro_nome nome do filtro personalizado (ex: 'CONTROLE ATUALIZADO', 'CONTROLE 2026')
    '''
    if not email:
        email = os.getenv("EMAIL")
    if not senha:
        senha = os.getenv("SENHA")

    nav, wait = init_browser()

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

        time.sleep(2)
        nav.get('https://www.integra.adv.br/moderno/modulo/50/default.asp')
        log.info("Página de relatórios acessada")

        time.sleep(2)

        # Seleciona filtro fase - Todos
        filtro_fase = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div[2]/div[21]/div[1]/button')))
        scroll_to_element(nav, filtro_fase)
        nav.execute_script("arguments[0].click();", filtro_fase)
        log.info("Filtro de 'Fase' aberto")

        todos_opt = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div[2]/div[21]/div[1]/div/div/ul/li[1]/a/span[2]')))
        todos_opt.click()
        log.info("Filtro 'Todos' clicado")

        # Filtro status - Ativo ou Inativo
        filtro_status = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div[2]/div[13]/div[2]/button')))
        scroll_to_element(nav, filtro_status)
        nav.execute_script("arguments[0].click();", filtro_status)
        log.info("Filtro de 'status' aberto")

        if status == 'Ativo':
            status_opt = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div[2]/div[13]/div[2]/div/ul/li[2]/label')))
        elif status == 'Inativo':
            status_opt = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div[2]/div[13]/div[2]/div/ul/li[3]/label')))
        else:
            raise ScrapperError('O valor de status da def get_data deve ser "Ativo" ou "Inativo"')

        time.sleep(2)

        status_opt.click()
        log.info(f"Filtro 'status' selecionado como {status}")

        time.sleep(2)

        pesquisar_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnPesquisar"]')))
        scroll_to_element(nav, pesquisar_button)
        nav.execute_script("arguments[0].click();", pesquisar_button)
        log.info("Botão 'Pesquisar' clicado")

        time.sleep(8)
        excel_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frmRelatorio"]/div[3]/div[1]/div/ul/li[5]/div/span')))
        scroll_to_element(nav, excel_button)
        excel_button.click()
        log.info("Relatório Excel selecionado")

        time.sleep(2)

        modelo_slct = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/button')))
        scroll_to_element(nav, modelo_slct)
        modelo_slct.click()

        time.sleep(2)
        modelo_label = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/div[1]/div/div/input')))
        modelo_label.send_keys(filtro_nome)
        time.sleep(2)
        try:
            contr_atualizado = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/div[1]/ul/li[3]/label/span')))
            contr_atualizado.click()
        except Exception:
            contr_atualizado = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divRelatorioMovimentacao"]/div/div/div/div/div[1]/ul/li[2]/label')))
            contr_atualizado.click()
        log.info(f"Modelo de relatório '{filtro_nome}' selecionado")

        time.sleep(3)

        gerar_excel = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnIlimitado"]')))
        scroll_to_element(nav, gerar_excel)
        gerar_excel.click()
        log.info("Download do relatório iniciado")

        time.sleep(20)
        log.info('get_data finalizada com sucesso')

    except ScrapperError:
        raise
    except Exception as e:
        msg = f'Ocorreu erro na def get_data na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. ERROR: {e}'
        log.error(msg)
        raise ScrapperError(msg)
    
