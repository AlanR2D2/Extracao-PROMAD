import os
from datetime import datetime
import time
import shutil
import pandas as pd
import send_to_googlesheet
import traceback
from monitoring import log
import glob
from dotenv import load_dotenv

load_dotenv()

SHEET = os.getenv("sheet")  # ID da google sheet
DATA_FILTRO = os.getenv("DATA_FILTRO", "08/03/2024")  # Data mínima para filtro de distribuição

DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))


class FileError(Exception):
    """Exceção customizada para erros de arquivo — não mata o processo."""
    pass

def obter_pasta_downloads():
    """Retorna o diretório dedicado de downloads do Chrome, ou fallback para ~/Downloads."""
    import Extrai_Promad
    download_dir = Extrai_Promad.get_download_dir()
    if download_dir and os.path.exists(download_dir):
        log.info(f'Usando diretório de downloads dedicado: {download_dir}')
        return download_dir
    log.info('Obtendo pasta de downloads do sistema operacional...')
    return os.path.join(os.path.expanduser("~"), "Downloads")

def RenomeiaUltimoArq(nome, ext):
    log.info('Renomeando último arquivo baixado...')
    nome = nome.replace('/', '-')
    print(f'Renomeando arquivo baixado para {nome + ext}')

    time.sleep(10)

    pasta_downloads = obter_pasta_downloads()
    novo_nome = nome + '.' + ext
    novo_path = os.path.join(pasta_downloads, novo_nome)

    todos_arquivos = [os.path.join(pasta_downloads, f) for f in os.listdir(pasta_downloads)]
    if not todos_arquivos:
        raise FileError(f"Nenhum arquivo encontrado em {pasta_downloads}")
    path_ultimo_arq = max(todos_arquivos, key=os.path.getctime)

    print(f'Ultimo arquivo é {path_ultimo_arq}')

    if os.path.exists(novo_path):
        try:
            os.remove(novo_path)
        except Exception as e:
            msg = f'Erro ao tentar remover arquivo existente na linha {traceback.extract_tb(e.__traceback__)[0].lineno}: {e}'
            raise FileError(msg)

    for i in range(10):
        try:
            os.rename(path_ultimo_arq, novo_path)
            print(f'Arquivo renomeado para {novo_nome}')
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
        pasta_downloads = obter_pasta_downloads()
        src = os.path.join(pasta_downloads, arq)
        dst = os.path.join(DIR_SCRIPT, arq)

        existe_anterior = os.path.exists(dst)
        print(f'Existe anterior? {existe_anterior}')
        if existe_anterior:
            os.remove(dst)

        shutil.move(src, dst)
    except Exception as e:
        msg = f'Ocorreu um erro na Move_Down_to_dir (arq) na linha {traceback.extract_tb(e.__traceback__)[0].lineno}, ERROR: {e}'
        raise FileError(msg)

def arq_to_sheet(arq_ativo, arq_inativo, aba, sheet_id=None):
    log.info('Iniciando leitura de arquivos HTML (ativo e inativo) exportados como Excel e envio para aba no Google Sheets...')
    sheet_id = sheet_id or SHEET
    try:
        data_minima = pd.to_datetime(DATA_FILTRO, dayfirst=True)

        # DataFrame com status Ativo
        dfs_ativo = pd.read_html(arq_ativo, header=1)
        df_ativo = dfs_ativo[0].copy()
        df_ativo = df_ativo.astype(object).fillna('')
        df_ativo['Data Distribuição'] = pd.to_datetime(df_ativo['Data Distribuição'], errors='coerce', dayfirst=True)
        df_ativo = df_ativo[df_ativo['Data Distribuição'] > data_minima]
        df_ativo['Data Distribuição'] = df_ativo['Data Distribuição'].dt.strftime('%d/%m/%Y')
        df_ativo['Status'] = 'Ativo'

        # DataFrame com status Inativo
        if os.path.getsize(arq_inativo) > 0:
            dfs_inativo = pd.read_html(arq_inativo, header=1)
            df_inativo = dfs_inativo[0].copy()
            df_inativo = df_inativo.astype(object).fillna('')
            df_inativo['Data Distribuição'] = pd.to_datetime(df_inativo['Data Distribuição'], errors='coerce', dayfirst=True)
            df_inativo = df_inativo[df_inativo['Data Distribuição'] > data_minima]
            df_inativo['Data Distribuição'] = df_inativo['Data Distribuição'].dt.strftime('%d/%m/%Y')
            df_inativo['Status'] = 'Inativo'
            df = pd.concat([df_ativo, df_inativo], ignore_index=True)
        else:
            log.info("[INFO] Arquivo inativo com 0 bytes, enviando apenas dados de Ativo.")
            df = df_ativo.copy()

        coloumns = df.columns.to_list()
        values = df.values.tolist()

        send_to_googlesheet.EscreveValores(f'{aba}!A1', [coloumns], sheet_id)
        send_to_googlesheet.EscreveValores(f'{aba}!A2', values, sheet_id)

        log.info("[OK] Tabela combinada extraída e enviada com sucesso ao Google Sheets.")
        return df
    except Exception as e:
        linha = traceback.extract_tb(e.__traceback__)[0].lineno
        msg = f"[ERRO] Falha na função arq_to_sheet na linha {linha}: {e}"
        raise FileError(msg)


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
        raise FileError(msg)

def df_to_append_sheet(df, aba):
    log.info('Convertendo DataFrame para lista e fazendo append no Google Sheets...')
    try:
        df.fillna('', inplace=True)
        coloumns = df.columns.to_list()
        values = df.values.tolist()

        send_to_googlesheet.AppendLinhas(values, SHEET, aba)
    except Exception as e:
        msg = f'Erro na def df_to_append_sheet na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. Error: {e}'
        raise FileError(msg)

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
        raise FileError(msg)

def excluir_arquivos_xls(diretorio='.'):
    log.info('Iniciando exclusão de arquivos .xls no diretório...')
    arquivos_xls = glob.glob(os.path.join(diretorio, '*.xls'))

    for arquivo in arquivos_xls:
        try:
            os.remove(arquivo)
            print(f"[OK] Arquivo excluído: {arquivo}")
        except Exception as e:
            log.error(f"[ERRO] Não foi possível excluir {arquivo}: {e}")
