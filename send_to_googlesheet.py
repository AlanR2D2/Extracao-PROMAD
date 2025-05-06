import os.path
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from monitoring import Logger
import send_email
import traceback

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

log = Logger('PROMAD', silencer=False).get_logger()

def error (msg):
    print(msg)
    log.error (msg)
    send_email.send_email_error(msg)
    exit()


def Autentica():
    try:
        creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        return sheet
    except Exception as e:
        msg = f'Ocorreu erro na def Autentica do script send_to_google_sheet na linha {traceback.extract_tb(e1.__traceback__)[0].lineno}. ERROR: {e}'
        error(msg)

# Deve ser utilizada no seguinte formado: LerValores('Sigtrip!A1:B6')
def LerValores (intervalo, spreadsheetId):
    sheet = Autentica()
    valores = sheet.values().get(spreadsheetId=spreadsheetId, range=intervalo).execute()
    linhas = valores.get('values')
    
    # Verifica se há valores em 'linhas'
    if linhas:
        # Determina o número de colunas na planilha
        num_colunas = max(len(linha) for linha in linhas)
        
        # Preenche as células vazias com ''
        for linha in linhas:
            while len(linha) < num_colunas:
                linha.append('')
    
    return linhas

def EscreveValores(intervalo, valores, spreadsheetId):
    sheet = Autentica()
    try:
        # Verificar e formatar os valores corretamente
        if not isinstance(valores, list):
            raise ValueError("Os valores devem ser uma lista de listas.")
        
        for item in valores:
            if not isinstance(item, list):
                raise ValueError("Cada item dentro de valores deve ser uma lista.")
        
        # Atualizar os valores na planilha
        response = sheet.values().update(
            spreadsheetId=spreadsheetId, 
            range=intervalo, 
            valueInputOption='USER_ENTERED', 
            body={'values': valores}
        ).execute()
        
        return response
    except Exception as e1:
        msg = f'Ocorreu erro na def EscreveValores do script send_to_google_sheet na linha {traceback.extract_tb(e1.__traceback__)[0].lineno}. ERROR: {e1}'
        error(msg)

def LimpaIntervalo(intervalo, spreadsheetId):
    try:
        sheet = Autentica()
        sheet.values().clear(spreadsheetId=spreadsheetId, range = intervalo).execute()
    except Exception as e2:
        msg = f'Ocorreu erro na def LimpaIntervalo do script send_to_google_sheet na linha {traceback.extract_tb(e2.__traceback__)[0].lineno}. ERROR: {e2}'
        error(msg)

def AppendLinhas(valores, spreadsheetId, nome_da_aba):
    sheet = Autentica()
    
    # Obter a última linha com dados na planilha
    ultima_linha = sheet.values().get(spreadsheetId=spreadsheetId, range=nome_da_aba).execute()
    ultima_linha_numero = len(ultima_linha.get('values', [])) + 1

    # Calcular o intervalo para adicionar as novas linhas
    novo_intervalo = f"{nome_da_aba}!A{ultima_linha_numero}"

    # Atualizar os valores na nova posição
    sheet.values().update(spreadsheetId=spreadsheetId, range=novo_intervalo, valueInputOption='USER_ENTERED', body={'values': valores}).execute()

    print(f"Linhas adicionadas com sucesso no intervalo {novo_intervalo}")

def SubstituiUltimaLinha(valores, spreadsheetId, nome_da_aba):
    sheet = Autentica()

    # Obter a última linha com dados na planilha
    ultima_linha = sheet.values().get(spreadsheetId=spreadsheetId, range=nome_da_aba).execute()
    ultima_linha_numero = len(ultima_linha.get('values', []))

    # Calcular o intervalo da última linha para substituição
    intervalo = f"{nome_da_aba}!A{ultima_linha_numero}"

    # Atualizar os valores na última linha
    response = sheet.values().update(
        spreadsheetId=spreadsheetId,
        range=intervalo,
        valueInputOption='USER_ENTERED',
        body={'values': valores}
    ).execute()

    print(f"Última linha substituída com sucesso no intervalo {intervalo}")

    return response

def ExcluiUltimasLinhas(nome_da_aba, spreadsheetId, num_linhas=15):
    sheet = Autentica()

    # Obter todas as linhas com dados na planilha
    todas_linhas = sheet.values().get(spreadsheetId=spreadsheetId, range=nome_da_aba).execute()
    total_linhas = len(todas_linhas.get('values', []))

    # Verificar se há pelo menos num_linhas para excluir
    if total_linhas <= num_linhas:
        intervalo = f"{nome_da_aba}!A1:Z{total_linhas}"
    else:
        intervalo = f"{nome_da_aba}!A{total_linhas - num_linhas + 1}:Z{total_linhas}"

    # Limpar as últimas linhas
    sheet.values().clear(spreadsheetId=spreadsheetId, range=intervalo).execute()

    print(f"As últimas {num_linhas} linhas foram excluídas com sucesso.")