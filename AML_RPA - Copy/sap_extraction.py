import subprocess
import time
from win32com.client import GetObject
from config import system_SAP, local_path, startDateFull, endDateFull,startDateYear, startDateMonth,control_spreadsheet_id,env,fileNameExel
from google_drive_operations import GoogleApi, Upload_File,TextToExcel,DeletaArquivo, ExelToTxt,send_log_to_sheets
from googleapiclient.discovery import build
import win32gui


def SAP_Extraction():
    try:
        
        hwnd = win32gui.FindWindow(None, "SAP Logon 800")  # Procura a janela pelo título
        if hwnd:
            subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])
            
        
        fileNameExel.clear() 
        #print(fileNameExel)
        subprocess.check_call([r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe', '-system=' + system_SAP]) 
        time.sleep(10)
        SapGuiAuto = GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine 
        connection = application.Children(0)
        session = connection.Children(0)
        session.findById("wnd[0]").iconify()
        
        def zse16():
            print("Começa: " + "zse16")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nzse16"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "marc"
            session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxtI2-LOW").text = "br12"
            session.findById("wnd[0]/usr/txtMAX_SEL").text = ""
            session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
            session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
            session.findById("wnd[0]").sendVKey (8)
            session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = local_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zse16.txt"
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            return 'zse16.txt'

        def Y_lad_65000280():        
            print("Comecar o Y_lad_65000280")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nY_LAD_65000280"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = "br12"
            session.findById("wnd[0]/usr/ctxtSP$00002-LOW").setFocus()
            session.findById("wnd[0]/usr/ctxtSP$00002-LOW").caretPosition = 4
            session.findById("wnd[0]").sendVKey (8)
            
            session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").select()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = local_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Y_LAD_65000280.txt"
            session.findById("wnd[1]/tbar[0]/btn[0]").press() 
            return 'Y_LAD_65000280.txt'

        def Mb51_261_262():
            print("start mb51_261-262")
            
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb51"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "br12"
            session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "261"
            session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = "262"
            session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = startDateFull
            session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = endDateFull
            
            session.findById("wnd[0]/usr/radRFLAT_L").setFocus()
            session.findById("wnd[0]/usr/radRFLAT_L").select()
            session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/brhey00"
            
            session.findById("wnd[0]/usr/ctxtALV_DEF").setFocus()
            session.findById("wnd[0]/usr/ctxtALV_DEF").caretPosition = 8
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            
            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[0]").select()
            session.findById("wnd[1]/tbar[0]/btn[20]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = local_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "mb51_261-262.xlsx"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            return 'mb51_261-262.xlsx'
        
        def Mb51_7():
            print("Comecar o Mb51_7")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb51"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "br12"
            session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "7*"
            session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = ""
            session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = startDateFull
            session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = endDateFull
            session.findById("wnd[0]/usr/radRFLAT_L").setFocus()
            session.findById("wnd[0]/usr/radRFLAT_L").select()
            session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/brhey00"
            session.findById("wnd[0]/usr/ctxtALV_DEF").setFocus()
            session.findById("wnd[0]/usr/ctxtALV_DEF").caretPosition = 8
            session.findById("wnd[0]").sendVKey (8)
            
            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[0]").select()
            session.findById("wnd[1]/tbar[0]/btn[20]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = local_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "mb51_7.xlsx"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
            session.findById("wnd[1]/tbar[0]/btn[0]").press()  

            return 'mb51_7.xlsx'
            
        def KKS5():
            print("KKS5")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nKKS5"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/subKV2000:SAPMKKS0:0200/chkKKS00-INCLF").selected = True
            session.findById("wnd[0]/usr/subKV2000:SAPMKKS0:0200/chkKKS00-INCLI").selected = True
            session.findById("wnd[0]/usr/subKV2000:SAPMKKS0:0200/chkKKS00-INCLP").selected = True
            session.findById("wnd[0]/usr/subKV2000:SAPMKKS0:0200/ctxtKKS00-WERKS").text = "BR12"
            session.findById("wnd[0]/usr/subKV2000:SAPMKKS0:0200/chkKKS00-INCLP").setFocus()
            session.findById("wnd[0]/usr/ctxtKKS00-POPER").text = startDateMonth
            session.findById("wnd[0]/usr/txtKKS00-GJAHR").text = startDateYear
            session.findById("wnd[0]/usr/txtKKS00-GJAHR").setFocus
            session.findById("wnd[0]/usr/txtKKS00-GJAHR").caretPosition = 4

            session.findById("wnd[0]/usr/radKKS00-AWVAL").select()
            session.findById("wnd[0]/usr/chkKKS00-BATCH").selected = True
            session.findById("wnd[0]/usr/chkKKS00-TESTL").selected = False
            session.findById("wnd[0]/usr/chkKKS00-LISTF").selected = True
            session.findById("wnd[0]/usr/chkKKS00-LISTF").setFocus()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/tabsKABA01_TBSTR/tabpDATE/ssubKABA01_SUBSN:SAPLKABA:0211/chkKABA01-STNOW").setFocus()
            session.findById("wnd[1]/usr/tabsKABA01_TBSTR/tabpDATE/ssubKABA01_SUBSN:SAPLKABA:0211/chkKABA01-STNOW").selected = True
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").text = "VM12"
            session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").caretPosition = 4
            session.findById("wnd[1]/tbar[0]/btn[6]").press()
            session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem ("PRIMM","Column2")
            session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem ("PRIMM","Column2")
            session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem ("PRIMM","Column2")
            session.findById("wnd[2]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/tbar[0]/btn[13]").press()
            session.findById("wnd[2]/tbar[0]/btn[0]").press()


        def SM37():
            print("SM37")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nSM37"
            session.findById("wnd[0]").sendVKey (0)

            session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = startDateFull
            session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").text = endDateFull
            session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").setFocus()
            session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").caretPosition = 7

            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(5)

            while True:
                try:
                    # Acessa o elemento que exibe o status e define o foco
                    sap_gui_screen = session.FindById("wnd[0]/usr/lbl[57,14]")
                    sap_gui_screen.SetFocus()

                    # Obtém o texto atual do elemento
                    texto_atual = sap_gui_screen.Text
                    print(f"Texto atual: {texto_atual}")

                    # Verifica se o texto é "Finished"
                    if texto_atual.strip() == "Finished":
                        print("O valor do campo é 'Finished'.")
                        break  # Sai do loop se o valor for 'Finished'
                    else:
                        print("O valor do campo não é 'Finished'.")
                    
                except Exception as e:
                    print(f"Ocorreu um erro: {e}")
                    break
                
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                # Pausa o loop por alguns segundos antes de verificar novamente
                time.sleep(90)  # Aguarda 90 segundos """

        def S_ALR_87013127():
            print("S_ALR_87013127")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nS_ALR_87013127"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxtRT_WERKS-LOW").text = "BR12"
            session.findById("wnd[0]/usr/ctxtRT_MATNR-LOW").text = ""
            session.findById("wnd[0]/usr/ctxtRT_AUART-LOW").text = "RM01"
            session.findById("wnd[0]/usr/txtP_PERIOV").text = startDateMonth
            session.findById("wnd[0]/usr/txtP_GJAHRV").text = startDateYear
            session.findById("wnd[0]/usr/txtP_PERIOB").text = startDateMonth
            session.findById("wnd[0]/usr/txtP_GJAHRB").text = startDateYear
            session.findById("wnd[0]/usr/ctxtP_VARS").text = "/BRNID02"
            session.findById("wnd[0]/usr/ctxtP_VARS").setFocus()
            session.findById("wnd[0]/usr/ctxtP_VARS").caretPosition = 8
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/mbar/menu[4]/menu[1]/menu[0]").select()
            session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[0]").select()
            session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-FORMAT").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[20]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = local_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "S_ALR_87013127.xlsx"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            return 'S_ALR_87013127.xlsx'
            
        
        def find_new_session(application, old_session_ids):
            for connection in application.Connections:
                if connection.Sessions.Count > len(old_session_ids):
                    for session in connection.Sessions:
                        if session.Id not in old_session_ids:
                            return session
            return None

        def KKBC_PKO():
            try:
            
                # Record the number of sessions before creating a new one
                old_session_ids = [session.Id for session in connection.Sessions]

                # Attempt to create a new session
                print("Attempting to open a new session...")
                connection.Children(0).CreateSession()
                
                # Wait for the new session to be created
                time.sleep(10)  # Increase sleep time if needed

                # Find the new session
                new_session = None
                for _ in range(10):  # Try several times to find the session
                    new_session = find_new_session(application, old_session_ids)
                    if new_session is not None:
                        break
                    time.sleep(2)  # Wait before trying again
                
                if new_session is None:
                    raise Exception("No new session found.")

                # Print session information
                print("Session Information:")
                try:
                    # Print session ID
                    print(f"Session ID: {new_session.Id}")
                    
                    # Verify if the session is ready by checking for a known element
                    new_session.findById("wnd[0]").text
                    print("Session is ready.")
                    
                    # Optional: Print session Title if it is available
                    # print(f"Session Title: {new_session.Title}")

                except Exception as e:
                    print(f"An error occurred while accessing session properties: {e}")

                # Example interaction
                print("Performing actions in the new session...")
                new_session.findById("wnd[0]/tbar[0]/okcd").text = "/nKKBC_PKO"
                new_session.findById("wnd[0]").sendVKey (0)
                new_session.findById("wnd[0]/usr/ctxtKKB0-WERKF").text = "BR12"
                new_session.findById("wnd[0]/usr/ctxtKKB0-WERKF").setFocus()
                new_session.findById("wnd[0]/usr/ctxtKKB0-WERKF").caretPosition = 4
                new_session.findById("wnd[0]/usr/radTIME_CUMU1").select()
                new_session.findById("wnd[0]/usr/txtD_PE1").text = startDateMonth
                new_session.findById("wnd[0]/usr/txtD_GJ1").text = startDateYear
                new_session.findById("wnd[0]/usr/txtD_PE2").text = startDateMonth
                new_session.findById("wnd[0]/usr/txtD_GJ2").text = startDateYear
                new_session.findById("wnd[0]/usr/txtD_GJ2").setFocus()
                new_session.findById("wnd[0]/usr/txtD_GJ2").caretPosition = 4
                new_session.findById("wnd[0]/mbar/menu[4]/menu[2]").select()
                new_session.findById("wnd[1]/usr/radOWAER").select()
                new_session.findById("wnd[1]/usr/radOWAER").setFocus()

            except Exception as e:
                print(f"An error occurred: {e}")
                    
        def Y_LAD_65000872():
            print("Y_LAD_65000872")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nY_LAD_65000872"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/txt$7K-KOKR").text = "CP00"
            session.findById("wnd[0]/usr/txt$7K-GJA1").text = startDateYear
            session.findById("wnd[0]/usr/txt$1GJAHR").text = startDateYear
            session.findById("wnd[0]/usr/txt$7K-PERV").text = startDateMonth
            session.findById("wnd[0]/usr/txt$7K-PERB").text = startDateMonth
            session.findById("wnd[0]/usr/txt$7K-GJA1").setFocus()
            session.findById("wnd[0]/usr/txt$7K-GJA1").caretPosition = 4
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/btn%__ZK-AUFN_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[23]").press()
            session.findById("wnd[2]/usr/ctxtDY_PATH").text = local_path
            session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "S_ALR_87013127.txt"
            session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 11
            session.findById("wnd[2]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/tbar[0]/btn[14]").press()   
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/mbar/menu[3]/menu[0]").select()
            session.findById("wnd[0]/tbar[1]/btn[14]").press()
            session.findById("wnd[1]/usr/ctxtLGRWO-OUT_FILE").setFocus()
            session.findById("wnd[1]/usr/ctxtLGRWO-OUT_FILE").caretPosition = 25
            session.findById("wnd[1]").sendVKey(4)
            session.findById("wnd[2]/usr/ctxtDY_PATH").text = local_path
            session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "Y_LAD_65000872.csv"
            session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 8
            session.findById("wnd[2]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/chkLGRWO-SAVE_PARMS").selected = True
            session.findById("wnd[1]/usr/chkLGRWO-SAVE_PARMS").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[2]/usr/btnSPOP-VAROPTION2").press()
            return "Y_LAD_65000872.csv"
            
            
        lista_de_metodos =[Mb51_7,Mb51_261_262,zse16,Y_lad_65000280,KKS5,SM37,S_ALR_87013127,KKBC_PKO,Y_LAD_65000872] 
        credits = GoogleApi()

        service = build('drive', 'v3', credentials=credits)
        #RUN SAP script
        for metodo in lista_de_metodos:
            i = metodo()
                        
            if i is not None:
                if ".txt" in i:
                    file = TextToExcel(i)
                    fileNameExel.append(file)
                    Upload_File(file,service)
                
                elif ".csv" in i:
                    file = TextToExcel(i)
                    fileNameExel.append(file)
                    Upload_File(file,service)
                    DeletaArquivo("S_ALR_87013127.txt")
               

                elif "S_ALR_87013127.xlsx" in i:
                    ExelToTxt(local_path,i)
                
                else:
                    fileNameExel.append(i)
                    Upload_File(i,service)
                    
                    """if "Y_LAD_65000872" in i:
                    
                    try:
                        # Close the new session
                        connection.CloseSession('ses[1]') 
                    except Exception as e:
                        print(f"An error occurred while closing the session: {e}") """
                                

                
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
            
            send_log_to_sheets(log_message, "Error", "AML", creds="",spreadsheet_id=control_spreadsheet_id,env=env)


        