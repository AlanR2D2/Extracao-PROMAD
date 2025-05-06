import Extrai_Promad
from monitoring import Logger, manter_apenas_ultimos_logs
import Trata_arquivos
import send_to_googlesheet
from datetime import datetime
import os
import send_email
import traceback

# Deixa somente os 10 últimos Logs
manter_apenas_ultimos_logs(quantidade=10)

# Exclui todos arquivos .xls
Trata_arquivos.excluir_arquivos_xls()

log = Logger('PROMAD', silencer=False).get_logger()

def error (msg):
    print(msg)
    log.error (msg)
    send_email.send_email_error(msg)
    exit()

try:
    # Gera data com hora atual
    def data_hora_formatada():
        return datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

    sheet = '1e9smWH4dVpTomORLW2s1bIRU-EY_7cyKAF0R4EehaDU'

    log.info('Iniciando main')
    Extrai_Promad.get_data()

    now = data_hora_formatada()

    arq = f'CONTROLE_ATUALIZADO_{now}'
    arq_ext = f'CONTROLE_ATUALIZADO_{now}.xls'

    Trata_arquivos.RenomeiaUltimoArq (arq, 'xls')
    Trata_arquivos.Move_Down_to_dir (arq_ext)
    send_to_googlesheet.LimpaIntervalo ('Dados!A2:Z', sheet)
    # Ignora dados antes de 08/03/2024
    Trata_arquivos.arq_to_sheet (arq_ext, 'Dados')
    os.remove(arq_ext)
except Exception as e:
    msg = f'Houve um erro na main na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. ERROR: {e}'
