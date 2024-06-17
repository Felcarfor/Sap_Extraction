import pandas as pd #Biblioteca para manipulação de dados
import win32com.client #Biblioteca para abrir programas do Windows(ex.: SAP)
import time
import subprocess
from dateutil.relativedelta import relativedelta
from gspread_pandas import Spread
from tkinter import Tk, Label, Entry, Button, font
from datetime import datetime, timedelta #Manipulação de Datas e Horários ( já do Python)
import tkinter as tk
import os
from oauth2client.service_account import ServiceAccountCredentials


# Define File and Folder for the output for Quotas report 
fileNameExel = ["mb51_261-262.XLSX","mb51_7.XLSX"]
folderdir = "" #"O:\\Shared drives\\teste\\"
spreadsheet_id = "" #'1d5VmlxbGGuHX5IEAt2n5rySWcxGkVrg5BCWEBgpHLFk'
mergeField = "" #'Row Labels'
mergeType = "" #'outer'
system_SAP = "" #"LAP"

dateToday = datetime.now().replace(day=1)
startDate = dateToday  - relativedelta(months=1)
endDate = dateToday - timedelta(days=1) 

startDate  = startDate.strftime('%m%d%Y')
endDate = endDate.strftime('%m%d%Y')
print(startDate) 
print(endDate)

print("Olá :)")
 
def SAP_Extraction():

    #Limpa a lista de nome de arquivos
    fileNameExel = []
    
    #Open SAP LOGON
    subprocess.check_call([r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe', '-system='+ system_SAP])
    time.sleep(10)
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
   # application.Visible = False
    connection = application.Children(0)
    session = connection.Children(0)
    
    """     WScript = win32com.client.Dispatch("{B54F3741-5B07-11cf-A4B0-00AA004A55E8}")
    WScript.ConnectObject(session, "on")
    WScript.ConnectObject(application, "on")
    """
    #MERGE POR  Material 
    def zse16():
        session.findById("wnd[0]/tbar[0]/okcd").text = "zse16"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "marc"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtI2-LOW").text = "br12"
        session.findById("wnd[0]/usr/ctxtI2-LOW").setFocus
        session.findById("wnd[0]/usr/ctxtI2-LOW").caretPosition = 4
        session.findById("wnd[0]/tbar[1]/btn[17]").press
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "BRhey00"
        session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
        session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
        session.findById("wnd[1]").sendVKey (8)
        session.findById("wnd[0]").sendVKey (8)
    
    def Y_lad_65000280():
        session.findById("wnd[0]/tbar[0]/okcd").text = "Y_lad_65000280"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/tbar[1]/btn[17]").press
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "brhey00"
        session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
        session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
        session.findById("wnd[1]").sendVKey (8)
        session.findById("wnd[0]").sendVKey (8)
        session.findById("wnd[1]").sendVKey (0)
        session.findById("wnd[1]").sendVKey (0)
    
    def Mb51_261_262():
        print("Comecar o mb51_261-262")
        session.findById("wnd[0]/tbar[0]/okcd").text = "mb51"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "br12"
        session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "262"
        session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = "262"
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = startDate
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = endDate
        session.findById("wnd[0]/usr/radRFLAT_L").setFocus()
        session.findById("wnd[0]/usr/radRFLAT_L").select()
        session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/brhey00"
        session.findById("wnd[0]/usr/ctxtALV_DEF").setFocus()
        session.findById("wnd[0]/usr/ctxtALV_DEF").caretPosition = 8
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        print("rodou o mb51_261-262")
        print("indo salvar como exel")
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderdir
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "mb51_261-262.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        print("salvou")
        
        return 'mb51_261-262.XLSX'
    
    def Mb51_7():
        print("Comecar o Mb51_7")
        session.findById("wnd[0]/tbar[0]/okcd").text = "mb51"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "br12"
        session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "7*"
        session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = ""
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = startDate
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = endDate
        session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/brhey00"
        session.findById("wnd[0]/usr/ctxtALV_DEF").setFocus()
        session.findById("wnd[0]/usr/ctxtALV_DEF").caretPosition = 8
        session.findById("wnd[0]").sendVKey (8)
        print("rodou o mb51_7")
        print("indo salvar como exel")
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderdir
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "mb51_7.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        print("salvou")
        return 'mb51_7.XLSX'
    
    lista_de_metodos = [zse16]
    
    #RUN SAP script
    for metodo in lista_de_metodos:
        name = metodo()
        fileNameExel.append(name) 
        print("retorna para a tela inicial do SAP")
        #Retorno para tela inicial do SAP 
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
        print('deu certo: ', name)
        
    # Close SAP GUI
    connection.CloseSession('ses[0]')     
   
   #Metodo para acessar as credenciais do google para poder editar o sheets
def LoadFromSheets(spreadsheet_id):
    # Carregue as credenciais da conta de serviço
    creds = {
  }
    # Crie um objeto Credentials a partir das credenciais
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds)

    s = Spread(spreadsheet_id,sheet=0,creds=credentials)

    return s
    
#Corre por uma lista de nome de arquivos e concatena eles, carrega o google sheets antigo, faz um merge com o antigo e novo, sobe para o sheets
def JoinAndSaveNewExtract():

    df = []
    for i in fileNameExel:  #Corre pelos Arquivos exel, aplica o "remove_leading_zeros", renomeia as colunas e faz o concat
        df.append(pd.read_excel(folderdir + i)) 
        
        #result1_concat = pd.merge([df], on='Material', how='')
        print(i)
    
    result1_concat = pd.concat(df, ignore_index=True)  

    s = LoadFromSheets(spreadsheet_id)
    gp = s
    worksheet = gp.sheet
    sheets_antigo_df = Spread.sheet_to_df(s)
    print(len(fileNameExel))
    
    if(len(fileNameExel) > 1):
        
        if sheets_antigo_df.empty:
            
            print("antigo vazio")
            
            #data_atualizacao = datetime.now().strftime('%d-%m-%Y %H:%M:%S')

            #merged_df['Data de Última Atualização'] = data_atualizacao
            #write the dataframe to the worksheet
            gp.df_to_sheet(result1_concat, sheet=worksheet,index=False, start='A1',replace=True) 

            
        else:
            print("Merge")
            merged_df = pd.merge(sheets_antigo_df,result1_concat, how='right', on='Material')
            
            # Obter a data e hora atual

            # create a gspread_pandas object and open the worksheet

            #write the dataframe to the worksheet
            gp.df_to_sheet(result1_concat, sheet=worksheet,index=False, start='A1',replace=True) 
            
        print("extraiu")
            
    else:
              
        if sheets_antigo_df.empty:
            
            print("antigo vazio")
            #write the dataframe to the worksheet
            gp.df_to_sheet(result1_concat, sheet=worksheet,index=False, start='A1',replace=True) 
            
        else: 
            merged_df = pd.merge(sheets_antigo_df,result1_concat, how='right', on='Material')
            data_atualizacao = datetime.now().strftime('%d-%m-%Y %H:%M:%S')        
            # Adicionar a nova coluna com a data de atualização
            merged_df['Data de Última Atualização'] = data_atualizacao
            #write the dataframe to the worksheet
            gp.df_to_sheet(result1_concat, sheet=worksheet,index=False, start='A1',replace=True) 
            # Obter a data e hora atual
             
        print("extraiu")
    
#Salva as informações   
def salvar_informacoes():
    global system_SAP, folderdir, spreadsheet_id, mergeField
    system_SAP = campo_sistema_sap.get()
    folderdir = campo_caminho_diretorio.get()
    spreadsheet_id = campo_spreadsheet_id.get()
    mergeField = campo_mergeField.get()

    with open("informacoes.txt", "w") as arquivo:
        arquivo.write("Sistema do SAP: {}\n".format(system_SAP))
        arquivo.write("Caminho do Diretório: {}\n".format(folderdir))
        arquivo.write("SpreadSheet ID: {}\n".format(spreadsheet_id))
        arquivo.write("MergeName: {}\n".format(mergeField))

    print("Informações salvas com sucesso!")

#Carrega informações
def carregar_informacoes():
    if os.path.exists("informacoes.txt"):
        with open("informacoes.txt", "r") as arquivo:
            linhas = arquivo.readlines()
            campo_sistema_sap.delete(0, tk.END)
            campo_sistema_sap.insert(0, linhas[0].split(": ")[1].strip())
            
            campo_caminho_diretorio.delete(0, tk.END)
            campo_caminho_diretorio.insert(0, linhas[1].split(": ")[1].strip())
            
            campo_spreadsheet_id.delete(0, tk.END)
            campo_spreadsheet_id.insert(0, linhas[2].split(": ")[1].strip())
            
            campo_mergeField.delete(0, tk.END)
            campo_mergeField.insert(0, linhas[3].split(": ")[1].strip())
            
            
            global system_SAP, folderdir, spreadsheet_id, mergeField
            system_SAP = campo_sistema_sap.get()
            folderdir = campo_caminho_diretorio.get()
            spreadsheet_id = campo_spreadsheet_id.get()
            mergeField = campo_mergeField.get()    
                        
            
    else:
        print("Arquivo de informações não encontrado.")

if __name__ == '__main__':
            
    # Verificar se o arquivo de informações existe e criá-lo se não existir
    if not os.path.exists("informacoes.txt"):
        with open("informacoes.txt", "w") as arquivo:
            arquivo.write("Sistema do SAP: "+ system_SAP + "\n")
            arquivo.write("Caminho do Diretório: "+ folderdir +"\n")
            arquivo.write("SpreadSheet ID: "+ spreadsheet_id + "\n")
            arquivo.write("Campo de Merge: "+ mergeField +"\n")
            
     # Criar os campos de entrada
   
    window = Tk()
    window.geometry("400x400")
    window.title("Configurações")

    # Estilo para os títulos
    titulo_fonte = font.Font(family="Helvetica", size=14, weight="bold")

    # Criar os campos de entrada
    Label(window, text="Extração do SAP", font=titulo_fonte).pack(pady=10)

    Label(window, text="Sistema do SAP:", font=("Helvetica", 10)).pack()
    campo_sistema_sap = Entry(window, bg="#E8E8E8", font=("Helvetica", 10))
    campo_sistema_sap.pack(padx=10, pady=5)

    Label(window, text="Caminho do Diretório:", font=("Helvetica", 10)).pack()
    campo_caminho_diretorio = Entry(window, bg="#E8E8E8", font=("Helvetica", 10), width=30)
    campo_caminho_diretorio.pack(padx=10, pady=5)

    Label(window, text="SpreadSheet ID:", font=("Helvetica", 10)).pack()
    campo_spreadsheet_id = Entry(window, bg="#E8E8E8", font=("Helvetica", 10), width=40)
    campo_spreadsheet_id.pack(padx=10, pady=5)

    Label(window, text="Campo de Merge:", font=("Helvetica", 10)).pack()
    campo_mergeField = Entry(window, bg="#E8E8E8", font=("Helvetica", 10))
    campo_mergeField.pack(padx=10, pady=5)

    # Carregar informações se existirem
    carregar_informacoes()
  
    # Botão para salvar as informações
    botao_salvar = Button(window, text="Salvar Informações", command=salvar_informacoes, bg="#4CAF50", fg="white", font=("Helvetica", 10, "bold")).pack(pady=10)
    # Botões de funcionalidade
    Button(window, text="Login SAP", command=SAP_Extraction, bg="#008CBA", fg="white", font=("Helvetica", 10, "bold")).pack(pady=5)
    Button(window, text="Excel para Sheets", command=JoinAndSaveNewExtract, bg="#008CBA", fg="white", font=("Helvetica", 10, "bold")).pack(pady=5)

    window.mainloop()