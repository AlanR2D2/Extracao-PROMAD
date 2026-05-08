import Extrai_Promad
from monitoring import log, manter_apenas_ultimos_logs
import Trata_arquivos
import send_to_googlesheet
import send_to_sharepoint
from datetime import datetime
import os
import send_email
import traceback
from dotenv import load_dotenv
import pytz

load_dotenv()

# Configuração das empresas
EMPRESAS = [
    {
        "nome": "Dra. Consumidor",
        "email": os.getenv("EMAIL"),
        "senha": os.getenv("SENHA"),
        "sheet": os.getenv("sheet"),
        "filtro": "CONTROLE ATUALIZADO",
        "sp_folder": "Planilhas dash",
        "sp_filename": "Promad.xlsx",
        "sp_timestamp": "promad_timestamp.xlsx",
    },
    {
        "nome": "Aeroline",
        "email": os.getenv("EMAIL_AEROLINE"),
        "senha": os.getenv("SENHA_AEROLINE"),
        "sheet": os.getenv("sheet_AEROLINE"),
        "filtro": "CONTROLE 2026",
        "sp_folder": "Planilhas Aerolines",
        "sp_filename": "Promad_aeroline.xlsx",
        "sp_timestamp": "promad_aeroline_timestamp.xlsx",
    },
]

# Deixa somente os 5 últimos Logs
manter_apenas_ultimos_logs(quantidade=5)

# Exclui todos arquivos .xls
Trata_arquivos.excluir_arquivos_xls()

def data_hora_formatada():
    return datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

try:
    for empresa in EMPRESAS:
        nome = empresa["nome"]
        emp_email = empresa["email"]
        emp_senha = empresa["senha"]
        emp_sheet = empresa["sheet"]
        emp_filtro = empresa["filtro"]

        log.info(f"========== Iniciando extração para: {nome} ==========")
        now = data_hora_formatada()

        # ------- Extraindo processos com status Ativo -------
        Extrai_Promad.get_data('Ativo', email=emp_email, senha=emp_senha, filtro_nome=emp_filtro)

        arq_ativo = f'CONTROLE_ATUALIZADO_ativo_{nome}_{now}'
        arq_ext_ativo = f'{arq_ativo}.xls'
        Trata_arquivos.RenomeiaUltimoArq(arq_ativo, 'xls')
        Trata_arquivos.Move_Down_to_dir(arq_ext_ativo)

        # ------- Extraindo processos com status Inativo (tentativas, fica com o MAIOR arquivo) -------
        NUM_TENTATIVAS = 3
        candidatos = []  # [(path, size)]
        for i in range(1, NUM_TENTATIVAS + 1):
            try:
                log.info(f"Tentativa {i}/{NUM_TENTATIVAS} para extrair 'Inativo' - {nome}")
                Extrai_Promad.get_data('Inativo', email=emp_email, senha=emp_senha, filtro_nome=emp_filtro)

                tentativa_base = f'CONTROLE_ATUALIZADO_inativo_{nome}_{now}_try{i:02d}'
                tentativa_file = f'{tentativa_base}.xls'
                Trata_arquivos.RenomeiaUltimoArq(tentativa_base, 'xls')
                Trata_arquivos.Move_Down_to_dir(tentativa_file)

                if os.path.exists(tentativa_file):
                    size = os.path.getsize(tentativa_file)
                    log.info(f"Arquivo '{tentativa_file}' obtido com {size} bytes")
                    candidatos.append((tentativa_file, size))
                else:
                    log.warning(f"Arquivo '{tentativa_file}' não encontrado após mover.")
            except Exception as e:
                log.warning(f"Falha na tentativa {i} para {nome}: {e}")

        if not candidatos:
            raise RuntimeError(f"Não foi possível obter nenhum arquivo 'Inativo' válido para {nome} após as tentativas.")

        candidatos.sort(key=lambda x: x[1], reverse=True)
        melhor_arquivo_inativo, melhor_tamanho = candidatos[0]
        log.info(f"Selecionado o maior arquivo de 'Inativo' para {nome}: '{melhor_arquivo_inativo}' ({melhor_tamanho} bytes)")

        arq_inativo = f'CONTROLE_ATUALIZADO_inativo_{nome}_{now}'
        arq_ext_inativo = f'{arq_inativo}.xls'

        if os.path.abspath(melhor_arquivo_inativo) != os.path.abspath(arq_ext_inativo):
            if os.path.exists(arq_ext_inativo):
                os.remove(arq_ext_inativo)
            os.rename(melhor_arquivo_inativo, arq_ext_inativo)

        for path, _ in candidatos[1:]:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception as e:
                log.warning(f"Não foi possível remover arquivo de tentativa '{path}': {e}")

        # ------- Envia para a planilha e atualiza timestamp -------
        send_to_googlesheet.LimpaIntervalo('Dados!A2:Z', emp_sheet)
        df_resultado = Trata_arquivos.arq_to_sheet(arq_ext_ativo, arq_ext_inativo, 'Dados', sheet_id=emp_sheet)

        timezone_brasilia = pytz.timezone('America/Sao_Paulo')
        ultima_atualizacao = datetime.now(timezone_brasilia).strftime("%d/%m/%Y %H:%M:%S")
        send_to_googlesheet.EscreveValores('Ultima atualização!A2', [[ultima_atualizacao]], emp_sheet)

        # ------- Upload para SharePoint (DESATIVADO) -------
        # try:
        #     sp_folder = empresa["sp_folder"]
        #     sp_filename = empresa["sp_filename"]
        #     sp_timestamp = empresa["sp_timestamp"]
        #     dir_script = os.path.dirname(os.path.abspath(__file__))
        #     xlsx_path = os.path.join(dir_script, sp_filename)
        #     timestamp_path = os.path.join(dir_script, sp_timestamp)
        #
        #     # Upload planilha de dados
        #     df_resultado.to_excel(xlsx_path, index=False, engine='openpyxl', sheet_name='Sheet')
        #     send_to_sharepoint.upload_file(xlsx_path, sp_folder, sp_filename)
        #     if os.path.exists(xlsx_path):
        #         os.remove(xlsx_path)
        #
        #     # Upload planilha de timestamp
        #     import pandas as pd
        #     df_ts = pd.DataFrame({'Última Atualização': [ultima_atualizacao]})
        #     df_ts.to_excel(timestamp_path, index=False, engine='openpyxl', sheet_name='Sheet')
        #     send_to_sharepoint.upload_file(timestamp_path, sp_folder, sp_timestamp)
        #     if os.path.exists(timestamp_path):
        #         os.remove(timestamp_path)
        # except Exception as sp_err:
        #     sp_msg = f'[SharePoint] Erro no upload para {nome}: {sp_err}'
        #     log.error(sp_msg)
        #     send_email.send_email_error(sp_msg)
        #     # Limpa arquivos temporários mesmo com erro
        #     for tmp in [xlsx_path, timestamp_path]:
        #         if os.path.exists(tmp):
        #             os.remove(tmp)

        # Remove os .xls trabalhados
        if os.path.exists(arq_ext_ativo):
            os.remove(arq_ext_ativo)
        if os.path.exists(arq_ext_inativo):
            os.remove(arq_ext_inativo)

        log.info(f"========== Extração para {nome} concluída ==========")

except Exception as e:
    msg = f'Houve um erro na main na linha {traceback.extract_tb(e.__traceback__)[0].lineno}. ERROR: {e}'
    log.error(msg)
    send_email.send_email_error(msg)
finally:
    Extrai_Promad.quit_browser()
