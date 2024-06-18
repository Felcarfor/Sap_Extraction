# Ferramenta de Extração e Integração de Dados SAP

## Visão Geral

Esta aplicação em Python facilita a extração de dados do SAP usando consultas pré-definidas (`mb51`) e integra os dados extraídos em um documento do Google Sheets.

## Requisitos

- Python 3.x
- Bibliotecas:
  - `pandas`: Para manipulação e análise de dados.
  - `win32com`: Automação do Windows para interagir com o SAP GUI.
  - `gspread_pandas`: Integração entre pandas e Google Sheets para atualização de dados.
  - `tkinter`: Biblioteca gráfica para interface de usuário.
  - `oauth2client`: Autenticação OAuth2 para acessar serviços do Google, essencial para Google Sheets.

## Configuração

### Instalação

Certifique-se de ter Python e as bibliotecas necessárias instaladas:
```bash
pip install pandas gspread-pandas pywin32 oauth2client
