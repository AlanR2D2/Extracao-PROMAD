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
    """Retorna o diretório dedicado de downloads do Chrome.

    NÃO cai mais para ~/Downloads: aquela pasta guarda exports antigos e o
    fallback fazia o script pegar um arquivo de meses atrás e publicá-lo na
    planilha como se fosse atual. Se o diretório dedicado sumiu (ex.: a limpeza
    horária do /tmp apagou no meio da execução), é melhor falhar e deixar a
    tentativa seguinte refazer o download.
    """
    import Extrai_Promad
    download_dir = Extrai_Promad.get_download_dir()
    if download_dir and os.path.exists(download_dir):
        log.info(f'Usando diretório de downloads dedicado: {download_dir}')
        return download_dir
    raise FileError(
        f'Diretório dedicado de downloads indisponível (valor: {download_dir}). '
        'O download precisa ser refeito.'
    )

# Extensões que o Chrome usa enquanto o download ainda está em andamento
EXT_PARCIAIS = ('.crdownload', '.part', '.tmp')

# Tempo máximo de espera pelo término de um download (segundos)
TIMEOUT_DOWNLOAD = int(os.getenv("TIMEOUT_DOWNLOAD", "420"))


def aguardar_download_completo(pasta_downloads, timeout=TIMEOUT_DOWNLOAD, intervalo=2, estabilidade=3):
    """Espera o Chrome terminar o download e devolve o caminho do arquivo final.

    O download só é considerado concluído quando não há nenhum arquivo parcial
    (.crdownload) na pasta e o tamanho do arquivo mais recente para de crescer
    por `estabilidade` verificações seguidas. Sem isso o script pegava o
    .crdownload ainda em andamento e gerava arquivos de 0 bytes.
    """
    log.info(f'Aguardando conclusão do download em {pasta_downloads} (timeout={timeout}s)...')
    limite = time.time() + timeout
    ultimo_tamanho = -1
    repeticoes = 0

    while time.time() < limite:
        try:
            nomes = os.listdir(pasta_downloads)
        except FileNotFoundError:
            raise FileError(f"Pasta de downloads não existe: {pasta_downloads}")

        parciais = [n for n in nomes if n.lower().endswith(EXT_PARCIAIS)]
        completos = [
            os.path.join(pasta_downloads, n)
            for n in nomes
            if not n.lower().endswith(EXT_PARCIAIS)
        ]
        completos = [p for p in completos if os.path.isfile(p) and os.path.getsize(p) > 0]

        if parciais or not completos:
            # Ainda baixando (ou nada baixado ainda) — reinicia a contagem de estabilidade
            ultimo_tamanho = -1
            repeticoes = 0
            time.sleep(intervalo)
            continue

        candidato = max(completos, key=os.path.getctime)
        tamanho = os.path.getsize(candidato)

        if tamanho == ultimo_tamanho:
            repeticoes += 1
            if repeticoes >= estabilidade:
                log.info(f'Download concluído: {candidato} ({tamanho} bytes)')
                return candidato
        else:
            ultimo_tamanho = tamanho
            repeticoes = 0

        time.sleep(intervalo)

    raise FileError(
        f"Timeout de {timeout}s aguardando o download terminar em {pasta_downloads}. "
        f"Conteúdo atual: {os.listdir(pasta_downloads)}"
    )


def RenomeiaUltimoArq(nome, ext):
    log.info('Renomeando último arquivo baixado...')
    nome = nome.replace('/', '-')
    print(f'Renomeando arquivo baixado para {nome + ext}')

    pasta_downloads = obter_pasta_downloads()
    novo_nome = nome + '.' + ext
    novo_path = os.path.join(pasta_downloads, novo_nome)

    path_ultimo_arq = aguardar_download_completo(pasta_downloads)

    print(f'Ultimo arquivo é {path_ultimo_arq}')

    if os.path.exists(novo_path):
        try:
            os.remove(novo_path)
        except Exception as e:
            msg = f'Erro ao tentar remover arquivo existente na linha {traceback.extract_tb(e.__traceback__)[0].lineno}: {e}'
            raise FileError(msg)

    ultimo_erro = None
    for i in range(10):
        try:
            os.rename(path_ultimo_arq, novo_path)
            print(f'Arquivo renomeado para {novo_nome}')
            break
        except Exception as e:
            ultimo_erro = e
            print('Parece que o arquivo ainda não terminou seu download ou está aberto...')
            print(f'Erro: {e}')
            time.sleep(5)
    else:
        raise FileError(f'Não foi possível renomear {path_ultimo_arq} para {novo_path}: {ultimo_erro}')

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
