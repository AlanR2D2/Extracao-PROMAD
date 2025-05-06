import os
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
import glob
import sys

class Logger:
    def __init__(
        self,
        log_name="default",
        logger_name=None,
        silencer=None,
        log_level=logging.DEBUG,
    ):
        """
        :param log_name: Nome usado para a pasta de logs e, caso logger_name não seja fornecido,
                         também para o nome do logger.

        :param logger_name: Caso queira especificar manualmente o nome do logger,
                            diferente do nome da pasta de logs.

        :param silencer: Silencia retornos de bibliotecas indesejados, majoritariamente retornos de API.

        :param log_level: Defina o nível de loglevel que o usuário deseja, por padrão, utiliza DEBUG.
        """
        if logger_name is None:
            logger_name = log_name

        self.logger = logging.getLogger(logger_name)

        if not self.logger.hasHandlers():
            self.logger.setLevel(log_level)

        formatter = logging.Formatter(
            "%(asctime)s %(lineno)4d   %(funcName)-12s     %(message)s"
        )

        actual_datetimestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        current_dir = os.path.dirname(os.path.abspath(__file__))
        logs_main_folder = os.path.join(current_dir, "logs")
        logs_folder = os.path.join(logs_main_folder, log_name)
        os.makedirs(logs_folder, exist_ok=True)

        log_filename = datetime.now().strftime(f"{log_name}_{actual_datetimestamp}.log")
        log_path = os.path.join(logs_folder, log_filename)

        file_handler = RotatingFileHandler(
            log_path, maxBytes=10 * 1024 * 1024, backupCount=300, encoding="utf-8"
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)

        if silencer:
            for log in silencer:
                logging.getLogger(log).setLevel(logging.WARNING)

        # Sempre adiciona o console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter(
            "%(message)s"
        )
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)

    def get_logger(self):
        return self.logger
    
    def instance_log (self):
        log = self.get_logger()
        return log
    
def manter_apenas_ultimos_logs(caminho_logs='./logs/PROMAD', quantidade=5):
    try:
        # Verifica se o diretório existe
        if not os.path.exists(caminho_logs):
            print(f"[AVISO] A pasta {caminho_logs} não existe.")
            return

        # Lista todos os arquivos de log da pasta
        arquivos_log = glob.glob(os.path.join(caminho_logs, '*'))

        # Ordena os arquivos por data de modificação (do mais recente para o mais antigo)
        arquivos_log.sort(key=os.path.getmtime, reverse=True)

        # Mantém os N mais recentes
        arquivos_a_remover = arquivos_log[quantidade:]

        for arquivo in arquivos_a_remover:
            os.remove(arquivo)
            print(f"[OK] Removido: {arquivo}")

        print(f"Limpeza de arquivos log. {len(arquivos_log) - len(arquivos_a_remover)} arquivos mantidos.")

    except Exception as e:
        print(f"[ERRO] Falha ao limpar os logs: {e}")



log = Logger('PROMAD', silencer=['selenium']).get_logger()


