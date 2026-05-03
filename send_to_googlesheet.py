import os.path
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from monitoring import log
import traceback
import time

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

MAX_RETRIES = 3
RETRY_DELAY = 5  # segundos entre tentativas

# --- Cache de autenticação (singleton) ---
_cached_sheet = None


class SheetError(Exception):
    """Exceção customizada para erros do Google Sheets — não mata o processo."""
    pass


def Autentica():
    global _cached_sheet
    if _cached_sheet is not None:
        return _cached_sheet
    try:
        log.info("Autenticando com Google Sheets via 'credentials.json'")
        creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        _cached_sheet = service.spreadsheets()
        log.info("Autenticação realizada com sucesso")
        return _cached_sheet
    except Exception as e:
        msg = f"Ocorreu erro na def Autentica do script send_to_google_sheet na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. ERROR: {e}"
        log.error(msg)
        raise SheetError(msg)

def LerValores(intervalo, spreadsheetId):
    log.info(f"Lendo valores do intervalo '{intervalo}' na planilha '{spreadsheetId}'")
    sheet = Autentica()
    valores = sheet.values().get(spreadsheetId=spreadsheetId, range=intervalo).execute()
    linhas = valores.get('values')

    if linhas:
        num_colunas = max(len(linha) for linha in linhas)
        for linha in linhas:
            while len(linha) < num_colunas:
                linha.append('')

    log.info(f"{len(linhas)} linhas obtidas com sucesso do intervalo '{intervalo}'")
    return linhas

def EscreveValores(intervalo, valores, spreadsheetId):
    sheet = Autentica()
    try:
        log.info(f"Escrevendo valores no intervalo '{intervalo}' da planilha '{spreadsheetId}'")
        if not isinstance(valores, list):
            raise ValueError("Os valores devem ser uma lista de listas.")
        for item in valores:
            if not isinstance(item, list):
                raise ValueError("Cada item dentro de valores deve ser uma lista.")

        for tentativa in range(1, MAX_RETRIES + 1):
            try:
                response = sheet.values().update(
                    spreadsheetId=spreadsheetId, 
                    range=intervalo, 
                    valueInputOption='USER_ENTERED', 
                    body={'values': valores}
                ).execute()
                log.info(f"Valores atualizados com sucesso no intervalo '{intervalo}'")
                return response
            except Exception as retry_err:
                if tentativa < MAX_RETRIES:
                    log.warning(f"Tentativa {tentativa}/{MAX_RETRIES} falhou para EscreveValores: {retry_err}. Retentando em {RETRY_DELAY}s...")
                    time.sleep(RETRY_DELAY)
                else:
                    raise
    except Exception as e1:
        msg = f"Ocorreu erro na def EscreveValores do script send_to_google_sheet na linha {traceback.extract_tb(e1.__traceback__)[0].lineno}. ERROR: {e1}"
        log.error(msg)
        raise SheetError(msg)

def LimpaIntervalo(intervalo, spreadsheetId):
    try:
        log.info(f"Limpando intervalo '{intervalo}' da planilha '{spreadsheetId}'")
        sheet = Autentica()
        for tentativa in range(1, MAX_RETRIES + 1):
            try:
                sheet.values().clear(spreadsheetId=spreadsheetId, range=intervalo).execute()
                log.info(f"Intervalo '{intervalo}' limpo com sucesso")
                return
            except Exception as retry_err:
                if tentativa < MAX_RETRIES:
                    log.warning(f"Tentativa {tentativa}/{MAX_RETRIES} falhou para LimpaIntervalo: {retry_err}. Retentando em {RETRY_DELAY}s...")
                    time.sleep(RETRY_DELAY)
                else:
                    raise
    except Exception as e2:
        msg = f"Ocorreu erro na def LimpaIntervalo do script send_to_google_sheet na linha {traceback.extract_tb(e2.__traceback__)[0].lineno}. ERROR: {e2}"
        log.error(msg)
        raise SheetError(msg)

def AppendLinhas(valores, spreadsheetId, nome_da_aba):
    sheet = Autentica()
    log.info(f"Adicionando novas linhas na aba '{nome_da_aba}' da planilha '{spreadsheetId}'")

    ultima_linha = sheet.values().get(spreadsheetId=spreadsheetId, range=nome_da_aba).execute()
    ultima_linha_numero = len(ultima_linha.get('values', [])) + 1
    novo_intervalo = f"{nome_da_aba}!A{ultima_linha_numero}"

    sheet.values().update(
        spreadsheetId=spreadsheetId, 
        range=novo_intervalo, 
        valueInputOption='USER_ENTERED', 
        body={'values': valores}
    ).execute()

    log.info(f"Linhas adicionadas com sucesso no intervalo {novo_intervalo}")
    print(f"Linhas adicionadas com sucesso no intervalo {novo_intervalo}")

def SubstituiUltimaLinha(valores, spreadsheetId, nome_da_aba):
    sheet = Autentica()
    log.info(f"Substituindo última linha da aba '{nome_da_aba}' da planilha '{spreadsheetId}'")

    ultima_linha = sheet.values().get(spreadsheetId=spreadsheetId, range=nome_da_aba).execute()
    ultima_linha_numero = len(ultima_linha.get('values', []))
    intervalo = f"{nome_da_aba}!A{ultima_linha_numero}"

    response = sheet.values().update(
        spreadsheetId=spreadsheetId,
        range=intervalo,
        valueInputOption='USER_ENTERED',
        body={'values': valores}
    ).execute()

    log.info(f"Última linha substituída com sucesso no intervalo {intervalo}")
    print(f"Última linha substituída com sucesso no intervalo {intervalo}")
    return response

def ExcluiUltimasLinhas(nome_da_aba, spreadsheetId, num_linhas=15):
    sheet = Autentica()
    log.info(f"Excluindo as últimas {num_linhas} linhas da aba '{nome_da_aba}' da planilha '{spreadsheetId}'")

    todas_linhas = sheet.values().get(spreadsheetId=spreadsheetId, range=nome_da_aba).execute()
    total_linhas = len(todas_linhas.get('values', []))

    if total_linhas <= num_linhas:
        intervalo = f"{nome_da_aba}!A1:Z{total_linhas}"
    else:
        intervalo = f"{nome_da_aba}!A{total_linhas - num_linhas + 1}:Z{total_linhas}"

    sheet.values().clear(spreadsheetId=spreadsheetId, range=intervalo).execute()

    log.info(f"As últimas {num_linhas} linhas foram excluídas com sucesso do intervalo {intervalo}")
    print(f"As últimas {num_linhas} linhas foram excluídas com sucesso.")
