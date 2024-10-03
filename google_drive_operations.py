from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from config import service_account_info, local_path,project_name
from openpyxl import Workbook
from datetime import datetime
import io
import pandas as pd
import os
import re


def GoogleApi():#acessa a API
    try:
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file']
        return service_account.Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    except Exception as e:
        print(f"Erro ao autenticar Google API: {e}")
        return None


def getFileFromGoogle(creds,folder_id,fileNameExel):#TRabalha com os dados aqui
    try:
        credits = creds
        if credits is None:
            raise Exception("Falha na autenticação com a Google API")

        print(fileNameExel)
        print(folder_id)
        
        dataframes = load_all_xlsx_files(folder_id, fileNameExel, credits)
        return dataframes
                            
    except Exception as e:
        print(f"Erro ao unir e salvar a nova extração: {e}")


def delete_existing_files(file_name, service,folder_id): #deleta os arquivos no drive
    try:
        query = f"name='{file_name}' and '{folder_id}' in parents"
        response = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        files = response.get('files', [])
        for file in files:
            try:
                print(f"Deleting file: {file['name']} with ID: {file['id']}")
                service.files().delete(fileId=file['id']).execute()
            except Exception as e:
                print(f"Error occurred while deleting file: {e}")
    except Exception as e:
        print(f"Erro ao deletar arquivos existentes: {e}")

def Upload_File(file_name, creds,folder_id):#Sobe os arquivos para a pasta do Drive
    try:
        service = build('drive', 'v3', credentials=creds)
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        full_path = local_path + file_name
        print(folder_id)
        delete_existing_files(file_name, service,folder_id)

        file_metadata = {
            'name': file_name.split('/')[-1],
            'parents': [folder_id]
        }
        media = MediaFileUpload(full_path, mimetype=mime_type)

        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

        print(f"File uploaded successfully. File ID: {file.get('id')}")
    except Exception as e:
        print(f"Erro ao fazer upload do arquivo: {e}")

def list_xlsx_files_in_folder(folder_id, creds):#pegas os arquivos que estao na pasta do Drive
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        query = f"'{folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
        response = drive_service.files().list(q=query, fields='files(id, name)').execute()
        return response.get('files', [])
    except Exception as e:
        print(f"Erro ao listar arquivos XLSX na pasta: {e}")
        return []

def download_and_load_xlsx(file_id,file_name, creds):
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)
        if(file_name == "Y_LAD_65000550.xlsx"):
            return pd.read_excel(fh,skiprows=9)  
        else:
            return pd.read_excel(fh)
    except Exception as e:
        print(f"Erro ao baixar e carregar arquivo XLSX: {e}")
        return None

def load_all_xlsx_files(folder_id, file_list, creds):
    try:
         # Lista todos os arquivos '.xlsx' na pasta
        files_gdrive = list_xlsx_files_in_folder(folder_id, creds)
        filtered_files = [file for file in files_gdrive if file['name'] in file_list]
        dataframes = {}
        for filtered_file in filtered_files:
            try:
                print(f"Carregando arquivo: {filtered_file['name']} com ID: {filtered_file['id']}")
                df = download_and_load_xlsx(filtered_file['id'],filtered_file['name'], creds)
                dataframes[filtered_file['name']] = df
            except Exception as e:
                print(f"Erro ao processar o arquivo {filtered_file['name']} com ID {filtered_file['id']}: {e}")

        return dataframes
    except Exception as e:
        print(f"Erro ao carregar todos os arquivos XLSX: {e}")
        return {}
        
def get_next_empty_row(service, spreadsheet_id, column):
    """
    Encontra a próxima linha vazia na coluna especificada.
    
    :param service: O serviço Google Sheets.
    :param spreadsheet_id: ID da planilha do Google Sheets.
    :param column: Coluna onde a mensagem será escrita.
    :return: O número da próxima linha vazia.
    """
    range_name = f'Sheet1!{column}1:{column}'
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    
    return len(values) + 1 if values else 1

def send_log_to_sheets(log_message: str, message_type: str, spreadsheet_id: str,env) -> None:
    """
    Envia uma mensagem de log para a próxima linha vazia em uma planilha do Google Sheets.

    :param log_message: Mensagem de log a ser enviada.
    :param message_type: Tipo de mensagem (Info, Erro, etc.).
    :param system: Sistema relacionado à mensagem de log.
    :param creds: Credenciais da conta de serviço.
    :param spreadsheet_id: ID da planilha do Google Sheets.
    """

    creds = GoogleApi()
        
    service = build('sheets', 'v4', credentials=creds)
    
    # Encontra a próxima linha vazia
    next_empty_row = get_next_empty_row(service, spreadsheet_id, 'A')
    range_name = f'Sheet1!A{next_empty_row}:E{next_empty_row}'
    
    # Obter data e hora atuais
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # Dados a serem escritos na planilha
    values = [[now,project_name, message_type, log_message,env ]]
    body = {'values': values}
    
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    
    #print(f"{result.get('updatedCells')} células atualizadas na linha {next_empty_row}.")

        
def save_dataframe_to_google_sheet(sheet_name,range,df, creds,file_name,spreadsheet_id):
    sheets_service = build('sheets', 'v4', credentials=creds)
    
    values = [df.columns.values.tolist()] + df.values.tolist()
    body = {
        'values': values
    }
    range_name = f'{sheet_name}!{range}'
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()
    print("Extraiu: "+file_name)
    
    
    
def get_dataframe_from_google_sheet(sheet_name, range_name, creds, spreadsheet_id):
    # Cria o serviço da API Google Sheets
    sheets_service = build('sheets', 'v4', credentials=creds)
    
    # Define o intervalo para a leitura
    range_name = f'{sheet_name}!{range_name}'
    
    # Obtém os valores da planilha
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    
    # Extrai os valores
    values = result.get('values', [])
    
    if not values:
        print('Nenhum dado encontrado na planilha.')
        return pd.DataFrame()  # Retorna um DataFrame vazio se não houver dados
    
    # Converte os valores para um DataFrame do pandas
    df = pd.DataFrame(values[1:], columns=values[0])  # A primeira linha é assumida como cabeçalho
    return df    

def TextToExcel(fileName):
    # Expressão regular para identificar a palavra "Table" na primeira linha
    table_pattern = re.compile(r'^Table:', re.IGNORECASE) #OUTRA FORMA DE FAZER SERIA DE COLOCAR O NOME DA PRIMERA COLUNA E TIRAR TUDO ANTES

    # Abrir o arquivo txt
    with open(local_path + fileName, 'r', encoding='utf-8', errors="ignore") as file:
        # Criar um novo Workbook
        workbook = Workbook()
        sheet = workbook.active

        # Contador para controlar a linha atual no arquivo Excel
        row_number = 1

        # Flag para controlar se devemos pular as próximas 3 linhas
        skip_next_lines = False

        # Iterar sobre cada linha do arquivo
        for line in file:
            # Verificar se a linha contém a palavra "Table" na primeira linha do arquivo
            if re.match(table_pattern, line) and not skip_next_lines:
                skip_next_lines = True
                continue  # Pular para a próxima iteração do loop sem processar esta linha

            # Se skip_next_lines for True, pular as próximas 1 linhas
            if skip_next_lines:
                skip_next_lines = False
                for _ in range(1):  # Pular uma linhas
                    next(file, None)
                continue  # Pular para a próxima iteração do loop

            # Dividir cada linha em colunas separadas por tabulação (ou outro delimitador)
            columns = line.strip().split("\t")
            
            # Lista para armazenar os índices das colunas a serem removidas

            # Escrever cada valor nas colunas correspondentes
            for j, value in enumerate(columns):
                if value:  # Verifica se o valor não é vazio
                    sheet.cell(row=row_number, column=j+1).value = value

            row_number += 1

        # Ajustar automaticamente a largura das colunas
        for column_cells in sheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            sheet.column_dimensions[column_letter].width = adjusted_width

        # Salvar o workbook como um arquivo Excel
        workbook = workbook.save(local_path + fileName[:-4] +".xlsx")

        file.close()

    DeletaArquivo(fileName)
    
    return fileName[:-4] +".xlsx"

def ExelToTxt(path, name):
    try:
        # Tenta carregar o arquivo Excel em um DataFrame
        df = pd.read_excel(os.path.join(path, name), skiprows=3)
         # Selecionar a coluna 'Order'
        df = df[["Order"]]
        # Ordenar o DataFrame pela coluna 'Order'
        df = df.sort_values(by='Order')
        
        df.to_csv(path + name[:-5] + ".txt", sep='\t', index=False)
        
        return name[:-5] +".txt"
    
    except FileNotFoundError:
        print(f"Arquivo {os.path.join(path, name)} não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro ao tentar ler o arquivo: {e}")
        
def DeletaArquivo(fileName):
    if os.path.exists(local_path + fileName):
        os.remove(local_path + fileName)
        print(f"O arquivo {fileName} foi deletado com sucesso.")
    else:
        print(f"O arquivo {fileName} não foi encontrado.")
