import os
from datetime import datetime
import time
import shutil
import pandas as pd
import send_to_googlesheet
import send_email
import traceback
from monitoring import Logger
import glob

SHEET = '1e9smWH4dVpTomORLW2s1bIRU-EY_7cyKAF0R4EehaDU'

log = Logger('PROMAD', silencer=False).get_logger()

def error (msg):
    print(msg)
    log.error (msg)
    send_email.send_email_error(msg)
    exit()

dir_script = os.getcwd()

# Obtém a data e hora atual
data_hora_atual = datetime.now()

def obter_pasta_downloads():
        
    try:

        sistema_operacional = os.name

        if sistema_operacional == 'nt':  # Windows
            pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        elif sistema_operacional == 'posix':  # Linux ou macOS
            pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        else:
            pasta_downloads = None

        return pasta_downloads
    
    except Exception as e:
        msg = f'Ocorreu um erro na def obter_pasta_downloads na linha {traceback.extract_tb(e.__traceback__)[0].lineno}, ERROR: {e}'
        error(msg)


def RenomeiaUltimoArq(nome, ext):
    nome = nome.replace('/','-')
    print(f'Renomeando arquivo baixado para {nome + ext}')

    time.sleep(10) #Tempo necessário para terminar o download

    pastaDownloads = obter_pasta_downloads()

    novoNome = nome + '.' + ext

    os.chdir(pastaDownloads)
    # Lista todos os arquivos da pasta Downloads
    todosArquivos = os.listdir()
    # Pega o arquivo mais recente
    path_ultimo_arq = max([os.path.join(pastaDownloads, f) for f in todosArquivos], key=os.path.getctime)

    print(f'Ultimo arquivo é {path_ultimo_arq}')

    try:
        # Caso tenha um arquivo com o mesmo nome, será excluído
        os.remove(novoNome)
    except FileNotFoundError:
        pass
    except Exception as e:
        msg = f'Erro ao tentar remover arquivo existente na linha {traceback.extract_tb(e.__traceback__)[0].lineno}: {e}'
        error (msg)

    # Renomeia o arquivo mais recentemente baixado
    for i in range(10):
        try:
            os.rename(path_ultimo_arq, novoNome)
            print(f'Arquivo renomeado para {novoNome}')
            break
        except Exception as e:
            print('Parece que o arquivo ainda não terminou seu download ou está aberto...')
            print(f'Erro: {e}')
            time.sleep(5)

#def que verifica se a pasta do dia já existe, caso exista a exclui, caso não somente cria a nova
# Tirei essa def pois decidi por todos os arquivos baixados junto com o script
# def CriaPasta_dir_script(nome_pasta):
#     os.chdir(dir_script)
#     pastaExiste = os.access(nome_pasta, os.F_OK)
#     if pastaExiste:
#         pass
#     else:
#         os.mkdir(nome_pasta)


def Move_Down_to_dir (arq):
    
    print(f'Movendo arquivo {arq} para pasta do script')
    time.sleep(2)

    try:
        os.chdir(dir_script)
        pastaDownloads = obter_pasta_downloads()
        # Verifica se já foi baixado, renomeado e transferido para a pasta do script um arquivo com o mesmo nome
        existe_anterior = os.access(os.path.join(dir_script,arq),os.F_OK)
        print(f'Existe anterior? {existe_anterior}')
        #Se existir um arq com o mesmo nome na pasta ele é deletado
        if existe_anterior:
            os.remove(os.path.join(dir_script,arq)) # Remove o arquivo duplicado
            time.sleep(3)
            os.chdir(pastaDownloads) # Volta para a pasta de Downloads
            time.sleep(3)
            shutil.move(arq,dir_script) # Recorta o arquivo para a pasta do script
            os.chdir(dir_script)
        else:
            os.chdir(pastaDownloads)
            time.sleep(3)
            shutil.move(arq,dir_script)
            os.chdir(dir_script)
    except Exception as e:
        msg = f'Ocorreu um erro na Move_Down_to_dir (arq) na linha {traceback.extract_tb(e.__traceback__)[0].lineno}, ERROR: {e}'
        error(msg)

def arq_to_sheet(arq, aba):
    """
    Extrai a primeira tabela de um arquivo .xls com conteúdo HTML e envia os dados
    para a aba especificada no Google Sheets.

    Parâmetros:
    - arq (str): Caminho do arquivo .xls exportado via navegador (HTML disfarçado de Excel).
    - aba (str): Nome da aba da planilha onde os dados serão escritos.

    Observação:
    O arquivo não é um Excel real, mas sim HTML estruturado com dados tabulares.
    """
    try:
        # Lendo todas as tabelas HTML do arquivo
        dfs = pd.read_html(arq, header=1)
        # Pegando a primeira tabela encontrada
        df = dfs[0]

        df.fillna('', inplace=True)

        # Convertendo a coluna para datetime
        df['Data Distribuição'] = pd.to_datetime(df['Data Distribuição'], errors='coerce', dayfirst=True)

        # Filtrando apenas as datas após 08/03/2024
        df = df[df['Data Distribuição'] > pd.to_datetime('08/03/2024', dayfirst=True)]

        # Convertendo de volta para string no formato dd/mm/yyyy
        df['Data Distribuição'] = df['Data Distribuição'].dt.strftime('%d/%m/%Y')

        coloumns = df.columns.to_list()
        values = df.values.tolist()
        
        send_to_googlesheet.EscreveValores(f'{aba}!A1', [coloumns],SHEET )
        send_to_googlesheet.EscreveValores(f'{aba}!A2', values, SHEET)

        print("[OK] Tabela extraída e enviada com sucesso")

    except Exception as e:
        print(f"[ERRO] Falha na função arq_to_sheet na linha {traceback.extract_tb(e.__traceback__)[0].lineno}: {e}")


def arq_to_append_sheet(arq,aba):
    try:
        content = pd.read_excel(arq, header=1)
        content.fillna('', inplace=True)
        coloumns = content.columns.to_list()
        values = content.values.tolist()

        send_to_googlesheet.AppendLinhas(values,SHEET,aba)
    except Exception as e:
        msg = f'Erro na def arq_to_sheet na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. Error: {e}'
        error(msg)

def df_to_append_sheet(df,aba):
    try:
        content = df
        content.fillna('', inplace=True)
        coloumns = content.columns.to_list()
        values = content.values.tolist()

        send_to_googlesheet.AppendLinhas(values,SHEET,aba)
    except Exception as e:
        msg = f'Erro na def arq_to_sheet na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. Error: {e}'
        error(msg)  

def arq_to_sheet_ate_hoje(arq, aba):
    try:
        # Leitura do arquivo Excel
        content = pd.read_excel(arq, header=0)
        content.fillna('', inplace=True)

        print(f'\nColunas do arq baixado: {content.keys()}\n')

        # Conversão da coluna 'Data' para datetime
        content['Data'] = pd.to_datetime(content['Data'], errors='coerce')

        # Obtenção da data de hoje
        hoje = pd.Timestamp.today().normalize()  # Normalize para remover a hora

        # Filtragem das linhas com datas menores ou iguais a hoje
        filtered_content = content[content['Data'] <= hoje]

        # Conversão da coluna 'Data' de volta para string
        filtered_content['Data'] = filtered_content['Data'].dt.strftime('%d/%m/%Y')

        # Extração dos nomes das colunas e dos valores
        columns = filtered_content.columns.tolist()
        values = filtered_content.values.tolist()

        # Limpeza e escrita dos dados na planilha do Google Sheets
        send_to_googlesheet.LimpaIntervalo(f'{aba}!A:J', SHEET)
        send_to_googlesheet.EscreveValores(f'{aba}!A1', [columns], SHEET)
        send_to_googlesheet.EscreveValores(f'{aba}!A2', values, SHEET)
    
    except Exception as e:
        msg = f'Erro na função arq_to_sheet_ate_hoje na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. Error: {e}'
        error(msg)  

def excluir_arquivos_xls(diretorio='.'):
    """
    Exclui todos os arquivos com extensão .xls do diretório especificado.

    :param diretorio: Caminho da pasta onde os arquivos .xls serão excluídos (padrão: diretório atual).
    """
    arquivos_xls = glob.glob(os.path.join(diretorio, '*.xls'))

    for arquivo in arquivos_xls:
        try:
            os.remove(arquivo)
            print(f"[OK] Arquivo excluído: {arquivo}")
        except Exception as e:
            print(f"[ERRO] Não foi possível excluir {arquivo}: {e}")
