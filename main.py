import Extrai_Promad
from monitoring import log, manter_apenas_ultimos_logs
import Trata_arquivos
import send_to_googlesheet
from datetime import datetime
import os
import send_email
import traceback
from dotenv import load_dotenv
import pytz

load_dotenv()
sheet = os.getenv("sheet")  # ID da google sheet

# Deixa somente os 5 últimos Logs
manter_apenas_ultimos_logs(quantidade=5)

# Exclui todos arquivos .xls
Trata_arquivos.excluir_arquivos_xls()

def error(msg):
    print(msg)
    log.error(msg)
    send_email.send_email_error(msg)
    exit(1)

def data_hora_formatada():
    return datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

try:
    now = data_hora_formatada()

    # ------- Extraindo processos com status Ativo -------
    Extrai_Promad.get_data('Ativo')

    arq_ativo = f'CONTROLE_ATUALIZADO_ativo_{now}'
    arq_ext_ativo = f'{arq_ativo}.xls'
    Trata_arquivos.RenomeiaUltimoArq(arq_ativo, 'xls')
    Trata_arquivos.Move_Down_to_dir(arq_ext_ativo)

    # ------- Extraindo processos com status Inativo (10 tentativas, fica com o MAIOR arquivo) -------
    NUM_TENTATIVAS = 10
    candidatos = []  # [(path, size)]
    for i in range(1, NUM_TENTATIVAS + 1):
        try:
            log.info(f"Tentativa {i}/{NUM_TENTATIVAS} para extrair 'Inativo'")
            Extrai_Promad.get_data('Inativo')

            # Renomeia o último .xls baixado para um nome único por tentativa
            tentativa_base = f'CONTROLE_ATUALIZADO_inativo_{now}_try{i:02d}'
            tentativa_file = f'{tentativa_base}.xls'
            Trata_arquivos.RenomeiaUltimoArq(tentativa_base, 'xls')
            # Move para a pasta de trabalho (como você já faz com os demais)
            Trata_arquivos.Move_Down_to_dir(tentativa_file)

            # Mede o tamanho e guarda como candidato
            if os.path.exists(tentativa_file):
                size = os.path.getsize(tentativa_file)
                log.info(f"Arquivo '{tentativa_file}' obtido com {size} bytes")
                candidatos.append((tentativa_file, size))
            else:
                log.warning(f"Arquivo '{tentativa_file}' não encontrado após mover.")
        except Exception as e:
            log.warning(f"Falha na tentativa {i}: {e}")

    if not candidatos:
        error("Não foi possível obter nenhum arquivo 'Inativo' válido após as tentativas.")

    # Seleciona o maior
    candidatos.sort(key=lambda x: x[1], reverse=True)
    melhor_arquivo_inativo, melhor_tamanho = candidatos[0]
    log.info(f"Selecionado o maior arquivo de 'Inativo': '{melhor_arquivo_inativo}' ({melhor_tamanho} bytes)")

    # Renomeia o melhor arquivo para o padrão final e remove os demais candidatos
    arq_inativo = f'CONTROLE_ATUALIZADO_inativo_{now}'
    arq_ext_inativo = f'{arq_inativo}.xls'

    # Se o melhor já tiver esse nome, ok; senão, renomeia
    if os.path.abspath(melhor_arquivo_inativo) != os.path.abspath(arq_ext_inativo):
        # Se já existir um alvo com o nome final (por alguma razão), remove antes
        if os.path.exists(arq_ext_inativo):
            os.remove(arq_ext_inativo)
        os.rename(melhor_arquivo_inativo, arq_ext_inativo)

    # Limpa os outros arquivos de tentativa
    for path, _ in candidatos[1:]:
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception as e:
            log.warning(f"Não foi possível remover arquivo de tentativa '{path}': {e}")

    # ------- Envia para a planilha e atualiza timestamp -------
    send_to_googlesheet.LimpaIntervalo('Dados!A2:Z', sheet)
    # Ignora dados antes de 08/03/2024 dentro do seu Trata_arquivos.arq_to_sheet
    Trata_arquivos.arq_to_sheet(arq_ext_ativo, arq_ext_inativo, 'Dados')

    # Horário Brasília
    timezone_brasilia = pytz.timezone('America/Sao_Paulo')
    ultima_atualizacao = datetime.now(timezone_brasilia).strftime("%d/%m/%Y %H:%M:%S")
    send_to_googlesheet.EscreveValores('Ultima atualização!A2', [[ultima_atualizacao]], sheet)

    # Remove os .xls trabalhados
    if os.path.exists(arq_ext_ativo):
        os.remove(arq_ext_ativo)
    if os.path.exists(arq_ext_inativo):
        os.remove(arq_ext_inativo)

except Exception as e:
    msg = f'Houve um erro na main na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. ERROR: {e}'
    log.error(msg)
    send_email.send_email_error(msg)
    # não faz exit aqui para permitir que o log + e-mail fechem o ciclo
