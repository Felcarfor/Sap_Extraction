from config import control_spreadsheet_id,env,folder_id,spreadsheet_id,project_name
from google_drive_operations import GoogleApi,send_log_to_sheets
from JoinAndSaveExtract import EDA
from sap_extraction import SAP_Extraction
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
#sys.stdout = Logger("RPA_AML_log_terminal.txt")

def Main_code():
    try:
        creds = GoogleApi()

            
        log_message = "Inicio da extracao principal "
        
        send_log_to_sheets(log_message, "Start", control_spreadsheet_id, env)
        
        SAP_Extraction(folder_id,creds)
        EDA(creds,folder_id,spreadsheet_id)
        
        log_message = "Final da extracao principal "

        send_log_to_sheets(log_message, "Success", control_spreadsheet_id, env) 
    
    except Exception as e:
        print(f"Capturado erro: {e}")  # Mensagem de depuração
        logging.error(f"Erro na execução principal: {e}")

        log_message = f"Erro na execução principal: {e} "
    
        send_log_to_sheets(log_message, "Error", control_spreadsheet_id, env)

if __name__ == '__main__':

    Main_code()
