import Extrai_Promad
from monitoring import log, manter_apenas_ultimos_logs
import Trata_arquivos
import send_to_googlesheet
from datetime import datetime
import os
import send_email
import traceback
from dotenv import load_dotenv

load_dotenv()

sheet = os.getenv("sheet")  # ID da google sheet

# Deixa somente os 5 últimos Logs
manter_apenas_ultimos_logs(quantidade=5)


# Exclui todos arquivos .xls
Trata_arquivos.excluir_arquivos_xls()

def error (msg):
    print(msg)
    log.error (msg)
    send_email.send_email_error(msg)
    exit()

#
try:
    # Gera data com hora atual
    def data_hora_formatada():
        return datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    
    now = data_hora_formatada()

    # ------- Extraindo processos com status Ativo -------

    Extrai_Promad.get_data('Ativo')

    arq_ativo = f'CONTROLE_ATUALIZADO_ativo_{now}'
    arq_ext_ativo = f'CONTROLE_ATUALIZADO_ativo_{now}.xls'
    Trata_arquivos.RenomeiaUltimoArq (arq_ativo, 'xls')
    Trata_arquivos.Move_Down_to_dir (arq_ext_ativo)

    # ------- Extraindo processos com status Inativo -------

    Extrai_Promad.get_data('Inativo')

    arq_inativo = f'CONTROLE_ATUALIZADO_inativo_{now}'
    arq_ext_inativo = f'CONTROLE_ATUALIZADO_inativo_{now}.xls'
    Trata_arquivos.RenomeiaUltimoArq (arq_inativo, 'xls')
    Trata_arquivos.Move_Down_to_dir (arq_ext_inativo)

    send_to_googlesheet.LimpaIntervalo ('Dados!A2:Z', sheet)
    # Ignora dados antes de 08/03/2024
    Trata_arquivos.arq_to_sheet (arq_ext_ativo,arq_ext_inativo, 'Dados')
    ultima_atualizacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    send_to_googlesheet.EscreveValores('Ultima atualização!A2', ultima_atualizacao, sheet)  
    
    os.remove(arq_ext_ativo)
    os.remove(arq_ext_inativo)

except Exception as e:
    msg = f'Houve um erro na main na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. ERROR: {e}'
    log.error(msg)
    send_email.send_email_error(msg)


