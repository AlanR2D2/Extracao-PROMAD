import os
from datetime import datetime
import time
import shutil
import pandas as pd
import send_to_googlesheet
import send_email
import traceback
from monitoring import log
import glob

SHEET = '1e9smWH4dVpTomORLW2s1bIRU-EY_7cyKAF0R4EehaDU'

def error(msg):
    print(msg)
    log.error(msg)
    send_email.send_email_error(msg)
    exit()

dir_script = os.getcwd()
data_hora_atual = datetime.now()

def obter_pasta_downloads():
    log.info('Obtendo pasta de downloads do sistema operacional...')
    try:
        sistema_operacional = os.name

        if sistema_operacional == 'nt':
            pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        elif sistema_operacional == 'posix':
            pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        else:
            pasta_downloads = None

        return pasta_downloads
    except Exception as e:
        msg = f'Ocorreu um erro na def obter_pasta_downloads na linha {traceback.extract_tb(e.__traceback__)[0].lineno}, ERROR: {e}'
        error(msg)

def RenomeiaUltimoArq(nome, ext):
    log.info('Renomeando último arquivo baixado...')
    nome = nome.replace('/', '-')
    print(f'Renomeando arquivo baixado para {nome + ext}')

    time.sleep(10)

    pastaDownloads = obter_pasta_downloads()
    novoNome = nome + '.' + ext

    os.chdir(pastaDownloads)
    todosArquivos = os.listdir()
    path_ultimo_arq = max([os.path.join(pastaDownloads, f) for f in todosArquivos], key=os.path.getctime)

    print(f'Ultimo arquivo é {path_ultimo_arq}')

    try:
        os.remove(novoNome)
    except FileNotFoundError:
        pass
    except Exception as e:
        msg = f'Erro ao tentar remover arquivo existente na linha {traceback.extract_tb(e.__traceback__)[0].lineno}: {e}'
        error(msg)

    for i in range(10):
        try:
            os.rename(path_ultimo_arq, novoNome)
            print(f'Arquivo renomeado para {novoNome}')
            break
        except Exception as e:
            print('Parece que o arquivo ainda não terminou seu download ou está aberto...')
            print(f'Erro: {e}')
            time.sleep(5)

def Move_Down_to_dir(arq):
    log.info('Movendo arquivo da pasta de downloads para pasta do script...')
    print(f'Movendo arquivo {arq} para pasta do script')
    time.sleep(2)

    try:
        os.chdir(dir_script)
        pastaDownloads = obter_pasta_downloads()
        existe_anterior = os.access(os.path.join(dir_script, arq), os.F_OK)
        print(f'Existe anterior? {existe_anterior}')
        if existe_anterior:
            os.remove(os.path.join(dir_script, arq))
            time.sleep(3)
            os.chdir(pastaDownloads)
            time.sleep(3)
            shutil.move(arq, dir_script)
            os.chdir(dir_script)
        else:
            os.chdir(pastaDownloads)
            time.sleep(3)
            shutil.move(arq, dir_script)
            os.chdir(dir_script)
    except Exception as e:
        msg = f'Ocorreu um erro na Move_Down_to_dir (arq) na linha {traceback.extract_tb(e.__traceback__)[0].lineno}, ERROR: {e}'
        error(msg)

def arq_to_sheet(arq, aba):
    log.info('Iniciando leitura de arquivo HTML exportado como Excel e envio para aba no Google Sheets...')
    try:
        dfs = pd.read_html(arq, header=1)
        df = dfs[0]
        df.fillna('', inplace=True)
        df['Data Distribuição'] = pd.to_datetime(df['Data Distribuição'], errors='coerce', dayfirst=True)
        df = df[df['Data Distribuição'] > pd.to_datetime('08/03/2024', dayfirst=True)]
        df['Data Distribuição'] = df['Data Distribuição'].dt.strftime('%d/%m/%Y')

        coloumns = df.columns.to_list()
        values = df.values.tolist()

        send_to_googlesheet.EscreveValores(f'{aba}!A1', [coloumns], SHEET)
        send_to_googlesheet.EscreveValores(f'{aba}!A2', values, SHEET)

        print("[OK] Tabela extraída e enviada com sucesso")
    except Exception as e:
        print(f"[ERRO] Falha na função arq_to_sheet na linha {traceback.extract_tb(e.__traceback__)[0].lineno}: {e}")

def arq_to_append_sheet(arq, aba):
    log.info('Iniciando leitura e append de planilha Excel para aba do Google Sheets...')
    try:
        content = pd.read_excel(arq, header=1)
        content.fillna('', inplace=True)
        coloumns = content.columns.to_list()
        values = content.values.tolist()

        send_to_googlesheet.AppendLinhas(values, SHEET, aba)
    except Exception as e:
        msg = f'Erro na def arq_to_append_sheet na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. Error: {e}'
        error(msg)

def df_to_append_sheet(df, aba):
    log.info('Convertendo DataFrame para lista e fazendo append no Google Sheets...')
    try:
        df.fillna('', inplace=True)
        coloumns = df.columns.to_list()
        values = df.values.tolist()

        send_to_googlesheet.AppendLinhas(values, SHEET, aba)
    except Exception as e:
        msg = f'Erro na def df_to_append_sheet na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. Error: {e}'
        error(msg)

def arq_to_sheet_ate_hoje(arq, aba):
    log.info('Iniciando leitura de planilha até data de hoje e atualização no Google Sheets...')
    try:
        content = pd.read_excel(arq, header=0)
        content.fillna('', inplace=True)

        print(f'\nColunas do arq baixado: {content.keys()}\n')

        content['Data'] = pd.to_datetime(content['Data'], errors='coerce')
        hoje = pd.Timestamp.today().normalize()
        filtered_content = content[content['Data'] <= hoje]
        filtered_content['Data'] = filtered_content['Data'].dt.strftime('%d/%m/%Y')

        columns = filtered_content.columns.tolist()
        values = filtered_content.values.tolist()

        send_to_googlesheet.LimpaIntervalo(f'{aba}!A:J', SHEET)
        send_to_googlesheet.EscreveValores(f'{aba}!A1', [columns], SHEET)
        send_to_googlesheet.EscreveValores(f'{aba}!A2', values, SHEET)
    except Exception as e:
        msg = f'Erro na função arq_to_sheet_ate_hoje na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. Error: {e}'
        error(msg)

def excluir_arquivos_xls(diretorio='.'):
    log.info('Iniciando exclusão de arquivos .xls no diretório...')
    arquivos_xls = glob.glob(os.path.join(diretorio, '*.xls'))

    for arquivo in arquivos_xls:
        try:
            os.remove(arquivo)
            print(f"[OK] Arquivo excluído: {arquivo}")
        except Exception as e:
            print(f"[ERRO] Não foi possível excluir {arquivo}: {e}")
