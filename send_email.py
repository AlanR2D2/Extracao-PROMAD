import smtplib
from email.message import EmailMessage
import os
from dotenv import load_dotenv
from monitoring import log

load_dotenv()

EMAIL_ADDRESS = os.getenv("EMAIL")  # Seu email Gmail
EMAIL_APP_PASSWORD = os.getenv("SENHA_APP_GMAIL")  # Senha de app do Google

def send_email_error(body: str):
    try:
        log.info("✉️ Preparando envio de e-mail de erro...")

        # Criando a mensagem
        msg = EmailMessage()
        msg['Subject'] = '❌ Erro no Script Promad'
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = 'alangcchagas@gmail.com'
        msg.set_content(body)

        # Enviando o e-mail
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_APP_PASSWORD)
            smtp.send_message(msg)

        log.info("📬 E-mail enviado com sucesso!")
        print('E-mail enviado com sucesso!')

    except Exception as e:
        log.error(f"❗Erro ao tentar enviar e-mail: {e}")
        print(f"[ERRO] Falha ao enviar e-mail: {e}")
