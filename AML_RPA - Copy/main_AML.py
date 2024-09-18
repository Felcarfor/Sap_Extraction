from config import  folder_id, fileNameExel,spreadsheet_id,control_spreadsheet_id,env,dateToday
from sap_extraction import SAP_Extraction
from google_drive_operations import GoogleApi, load_all_xlsx_files, save_dataframe_to_google_sheet,send_log_to_sheets
import numpy as np
import logging
import sys
import time
# Configuring the logger

#logging.basicConfig(filename='RPA_AML_log.txt',level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


class Logger:
    def __init__(self, filename):
        self.terminal = sys.stdout
        self.log = open(filename, "a")

    def write(self, message):
        # Split the message into multiple lines and add timestamp to each line.
        for line in message.splitlines():
            if line:  # Ignore empty lines
                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                self.terminal.write(f"{timestamp} - {line}\n")
                self.log.write(f"{timestamp} - {line}\n")
    
    def flush(self):
        self.terminal.flush()
        self.log.flush()

# Redirect stdout to the custom Logger
sys.stdout = Logger("RPA_AML_log_terminal.txt")

def JoinAndSaveNewExtract(creds):#TRabalha com os dados aqui
    try:
        credits = creds
        if credits is None:
            raise Exception("Falha na autenticação com a Google API")

        print(fileNameExel)
        print(folder_id)
        
        dataframes = load_all_xlsx_files(folder_id, fileNameExel, credits)
        for file_name, df in dataframes.items():
            if "mb51_261-262.xlsx" == file_name:
                df = df.sort_values(by='Material')

                df = df.groupby('Material').agg({'Quantity': 'sum', 'Amt.in Loc.Cur.': 'sum'}).reset_index()
                save_dataframe_to_google_sheet("1 - Download SAP - Weekly", f'A2:C', df, credits,file_name,spreadsheet_id)
                
                
            if "mb51_7.xlsx" == file_name:
                df = df.sort_values(by='Material')
                df = df.groupby('Material').agg({'Quantity': 'sum', 'Amt.in Loc.Cur.': 'sum'}).reset_index()
                save_dataframe_to_google_sheet("1 - Download SAP - Weekly", f'F2:H', df, credits,file_name,spreadsheet_id)
                
            if "zse16.xlsx" == file_name:
                df = df.sort_values(by='Material')
                df = df.fillna('')
                df = df.rename(columns={"Plant": "Plnt", "MRP Controller": "MRPCn"})
                df = df[["Material", 'Plnt', 'MRPCn']]
                save_dataframe_to_google_sheet("0 - SKUs info", f'H3:L', df, credits,file_name,spreadsheet_id)
                
            if "Y_LAD_65000280.xlsx" == file_name:
                df = df.sort_values(by='Material')
                df = df.rename(columns={" Standard price": "Standard price", "   per": "per"})
                df = df[["Material", "Standard price", "per", "BUn"]]
                df['Standard price'] = df['Standard price'].str.replace(',', '').astype(float)
                df["per"] = df["per"].str.replace(',', '').astype(float)
                df["STD Cost"] = df["Standard price"] / df["per"]
                df = df.fillna('')
                save_dataframe_to_google_sheet("0 - SKUs info", f'A3:F', df, credits,file_name,spreadsheet_id)
                
            if "Y_LAD_65000872.xlsx" == file_name:
                df = df[["Target","Total cost","Target Qty","Total","Lead column",'In.Price V',"Qty Varian","Res-Usage","RemInVar","OutPrice","Lote size","RemOutVar"]]  
                df = df.replace({float('NaN'): None})
                save_dataframe_to_google_sheet("5 - Variance Integration", f'A2:L', df, credits,file_name,spreadsheet_id)
                #FAZER O GROUP BY DOS VALORES POR MATERIAL E COLOCAR EM UMA TABELA OS MAIORES VER COM AS MENINAS A FORMULA PARA ISSO
                df = df.groupby('Lead column').agg({'Target': 'sum', 'Total cost': 'sum','Total': 'sum',}).reset_index()
                df = df.assign(Date=dateToday.strftime("%d/%m/%Y"))
                save_dataframe_to_google_sheet("Variance Sum", f'A1:L', df, credits,file_name,spreadsheet_id)


    
            
        return credits
    except Exception as e:
        logging.error(f"Erro ao unir e salvar a nova extração: {e}")
        log_message = f"Erro ao unir e salvar a nova extração: {e}"
        send_log_to_sheets(log_message, "Error", "AML", creds, control_spreadsheet_id, env)
        


if __name__ == '__main__':

    creds = GoogleApi()
    
    try:
        log_message = "Inicio da extracao principal"
        send_log_to_sheets(log_message, "Start", "AML", creds, control_spreadsheet_id, env)

        #SAP_Extraction()
        JoinAndSaveNewExtract(creds)
        
        log_message = "Sucesso na execucao principal"
        send_log_to_sheets(log_message, "Success", "AML", creds, control_spreadsheet_id, env)

    except Exception as e:
            print(f"Capturado erro: {e}")  # Mensagem de depuração
            logging.error(f"Erro na execução principal: {e}")

            log_message = f"Erro na execução principal: {e}"
            
            send_log_to_sheets(log_message, "Error", "AML", creds, control_spreadsheet_id, env)