import smtplib
from email.message import EmailMessage
import os
from dotenv import load_dotenv

load_dotenv()

EMAIL_ADDRESS = os.getenv("EMAIL") #Seu email gmail
EMAIL_APP_PASSWORD = os.getenv("SENHA_APP_GMAIL") #Senha app do google. Para gerar pesquise no google: Criar senha app

def send_email_error(body:str):
    # Criando a mensagem
    msg = EmailMessage()
    msg['Subject'] = 'Erro no Script promad'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = 'alangcchagas@gmail.com'
    msg.set_content('Houve falha.')

    # Enviando o e-mail
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_APP_PASSWORD)
        smtp.send_message(msg)

    print('E-mail enviado com sucesso!')