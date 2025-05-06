# 📊 Promad Data Automation

Este projeto automatiza a extração de relatórios do sistema jurídico **Promad**, realiza o tratamento do arquivo exportado (que não é um Excel padrão, mas sim uma estrutura HTML encapsulada), e envia os dados diretamente para uma aba específica em uma **planilha do Google Sheets**.

---

## 🚀 Funcionalidades

- Login automatizado no Promad via Selenium.
- Acesso à tela de relatórios e geração do relatório "📋 CONTROLE ATUALIZADO".
- Download automático do arquivo `.xls` (com estrutura HTML interna).
- Desbloqueio do arquivo baixado.
- Leitura do arquivo utilizando `pandas.read_html()` a partir do arquivo `sheet001.htm`.
- Envio dos dados extraídos para o **Google Sheets**.
- Registro de logs e envio de notificações por e-mail em caso de falhas.

---

## 🧱 Estrutura do Projeto

```text
📁 Promad/
├── main.py                  # Script principal que orquestra todo o processo
├── Extrai_Promad.py         # Realiza a automação do Promad com Selenium
├── Trata_arquivos.py        # Manipula arquivos baixados e extrai tabelas do HTML
├── send_to_googlesheet.py   # Lógica para autenticação e escrita no Google Sheets
├── send_email.py            # Envio de e-mails em caso de falha
├── monitoring.py            # Logger customizado com rotação de arquivos
├── .env                     # Credenciais sensíveis (não versionado)
├── credentials.json         # Credenciais de serviço do Google API
├── logs/
│   └── PROMAD/...           # Logs gerados por execução
└── downloads/               # (opcional) Pasta para onde os arquivos são movidos


