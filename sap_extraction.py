import subprocess
import time
from win32com.client import GetObject
from config import system_SAP, local_path, startDateFull, endDateFull,startDateYear, startDateMonth,fileNameExel,startDateFull2,endDateFull2
from google_drive_operations import Upload_File,DeletaArquivo,get_dataframe_from_google_sheet
import win32gui




def SAP_Extraction(folder_id,creds):
    try:
        while True:
            try:
                # Procura a janela pelo título
                hwnd = win32gui.FindWindow(None, "SAP Logon 800")
                if hwnd:
                    subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])

                # Conecta ao SAP GUI
                subprocess.check_call([r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe', '-system=' + system_SAP])
                time.sleep(10)
                
                # Obtém o objeto SAPGUI
                SapGuiAuto = GetObject('SAPGUI')
                # Obtém a aplicação SAP
                application = SapGuiAuto.GetScriptingEngine
                connection = application.Children(0)  # Altere o índice da conexão conforme necessário
                session = connection.Children(0)  # Altere o índice da sessão conforme necessário
                
                # Manipula a aplicação SAP conforme necessário
                fileNameExel.clear() # Certifique-se de definir 'fileNameExel' conforme apropriado

                print("Conexão ao SAP realizada com sucesso!")
                break  # Encerra o loop se a conexão for bem-sucedida

            except Exception as e:
                print(f"Ocorreu um erro ao tentar conectar ao SAP: {str(e)}")
                print("Tentando novamente em 5 segundos...")
                time.sleep(5)  # Intervalo antes de tentar novamente
                
        
        def md04(material):
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nmd04"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").text = material# loop aqui
            session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = "BR12"
            session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").setFocus()
            session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").caretPosition = 4
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/btnBUTTON_EZ_PS").press()
            session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0770/tabsPS_TAB/tabpPS_W").select()
            session.findById("wnd[0]/mbar/menu[0]/menu[4]").select()
            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[0]").select()
            session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME").caretPosition = 22
            session.findById("wnd[1]/tbar[0]/btn[20]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = local_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "md04_"+ str(material) +".xlsx"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 27
            session.findById("wnd[1]").sendVKey (0)
            return "md04_"+ str(material) +".xlsx"
            

        lista_de_metodos = [md04]

        df = get_dataframe_from_google_sheet("Sheets","A:A",creds,"1E4UGPw7lV73b8GPxLfPNxNtS9hGb_lbPPNCLSjQ80Qs")#passando o ID expeficio da onde tirar os materias

        serie = df["Material"]
        
        #RUN SAP script
        for item in serie:
            for metodo in lista_de_metodos:
                
                i = metodo(item)
                print("aa")        
                fileNameExel.append(i)
                Upload_File(i,creds,folder_id)
                DeletaArquivo(i)
                
                print("retorna para a tela inicial do SAP")
            
                
        # Close SAP GUI
        connection.CloseSession('ses[0]') 
        hwnd = win32gui.FindWindow(None, "SAP Logon 800")  # Procura a janela pelo título
        if hwnd:
            subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])
            
        print("Finalizou a extração")  

        return fileNameExel
    
    except Exception as e:
            print(f"Capturado erro: {e}")  # Mensagem de depuração
            log_message = f"Erro na execução principal: {e}"
            
           # send_log_to_sheets(log_message, "Error", "AML", creds="",spreadsheet_id=control_spreadsheet_id,env=env)


        