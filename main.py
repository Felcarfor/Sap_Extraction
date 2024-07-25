import pandas as pd #Biblioteca para manipulação de dados
import time
import subprocess
from dateutil.relativedelta import relativedelta
from gspread_pandas import Spread
from tkinter import Tk, Label, Entry, Button, font
from datetime import datetime, timedelta #Manipulação de Datas e Horários ( já do Python)
import tkinter as tk
import os
from oauth2client.service_account import ServiceAccountCredentials
import string
import json
import re
from openpyxl import Workbook
import pythoncom
pythoncom.CoInitialize()
from win32com.client import GetObject

# Define File and Folder for the output for Quotas report 
fileNameExel = ["mb51_261-262.XLSX","mb51_7.XLSX","zse16.XLSX","Y_LAD_65000280.XLSX"]
directory = "" #"O:\\Shared drives\\teste\\"
spreadsheet_id = "" #'1d5VmlxbGGuHX5IEAt2n5rySWcxGkVrg5BCWEBgpHLFk'
mergeField = "" #'Row Labels'
system_SAP = "" #"LAP"


dateToday = datetime.now()
startDate = dateToday.replace(day=1)
endDate = startDate - timedelta(days=1) 
endDate += relativedelta(months=1) 

startDate  = startDate.strftime('%m%d%Y')
endDate = endDate.strftime('%m%d%Y')
print(startDate) 
print(endDate)
print("v1.0")
 
def SAP_Extraction():

#Limpa a lista de nome de arquivos
    fileNameExel = []
    print(fileNameExel)
    subprocess.check_call([r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe', '-system=' + system_SAP]) 
    time.sleep(40)
    SapGuiAuto = GetObject('SAPGUI')
    
    application = SapGuiAuto.GetScriptingEngine 
    connection = application.Children(0)
    session = connection.Children(0)

    def zse16():
        print("Começa: " + "zse16")
        session.findById("wnd[0]/tbar[0]/okcd").text = "zse16"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "marc"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtI2-LOW").text = "br12"
        session.findById("wnd[0]/usr/txtMAX_SEL").text = ""
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        session.findById("wnd[0]").sendVKey (8)
        session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = directory
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "zse16.txt"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        return 'zse16.txt',5

    def Y_lad_65000280():        
        print("Comecar o Y_lad_65000280")
        session.findById("wnd[0]/tbar[0]/okcd").text = "Y_LAD_65000280"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = "br12"
        session.findById("wnd[0]/usr/ctxtSP$00002-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtSP$00002-LOW").caretPosition = 4
        session.findById("wnd[0]").sendVKey (8)
        
        session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = directory
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Y_LAD_65000280.txt"
        session.findById("wnd[1]/tbar[0]/btn[0]").press() 
        return 'Y_LAD_65000280.txt',3

    def Y_LAD_65000550():  
        print("Comecar o Y_LAD_65000550")
        session.findById("wnd[0]/tbar[0]/okcd").text = "Y_LAD_65000550"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "br12"
        session.findById("wnd[0]/usr/ctxtS_DISPO-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtS_CWERK-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").text = startDate
        session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").text = endDate
        session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").caretPosition = 9
        session.findById("wnd[0]").sendVKey (8)
        session.findById("wnd[0]/usr/lbl[27,16]").setFocus()
        session.findById("wnd[0]/usr/lbl[27,16]").caretPosition = 33
        session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = directory
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Y_LAD_65000550.txt"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        return 'Y_LAD_65000550.txt'
    
    def Mb51_261_262():
        print("Comecar o mb51_261-262")
        session.findById("wnd[0]/tbar[0]/okcd").text = "mb51"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "br12"
        session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "261"
        session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = "262"
        session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = startDate
        session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = endDate
        session.findById("wnd[0]/usr/radRFLAT_L").setFocus()
        session.findById("wnd[0]/usr/radRFLAT_L").select()
        session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/brhey00"
        session.findById("wnd[0]/usr/ctxtALV_DEF").setFocus()
        session.findById("wnd[0]/usr/ctxtALV_DEF").caretPosition = 8
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[0]").select()
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = directory
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "mb51_261-262.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        return 'mb51_261-262.XLSX',4
    
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
        
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[0]").select()
        session.findById("wnd[1]/tbar[0]/btn[20]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = directory
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "mb51_7.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        print("salvou")
        return 'mb51_7.XLSX',3
    lista_de_metodos = [Mb51_7,Mb51_261_262,zse16,Y_lad_65000280]
      
    #RUN SAP script
    for metodo in lista_de_metodos:
        i = metodo()
        
        if ".txt" in i[0]:
           fileNameExel.append(TextToExcel(i[0]))
        else:
            fileNameExel.append(i[0])

        print("retorna para a tela inicial do SAP")

        print(fileNameExel)
        #Retorno para tela inicial do SAP 
        for t in range(i[1]):
            session.findById("wnd[0]/tbar[0]/btn[12]").press()
        
        print('deu certo: ', i[0])
        
    # Close SAP GUI
    connection.CloseSession('ses[0]') 
     
    print("Finalizou a extração")  
   
   #Metodo para acessar as credenciais do google para poder editar o sheets
def LoadFromSheets(spreadsheet_id):
    # Carregue as credenciais da conta de serviço

    # Crie um objeto Credentials a partir das credenciais
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds)
    print("autenticou")
    return Spread(spreadsheet_id,sheet=0,creds=credentials)
    
    
def TextToExcel(fileName):
    # Expressão regular para identificar a palavra "Table" na primeira linha
    table_pattern = re.compile(r'^Table:', re.IGNORECASE) #OUTRA FORMA DE FAZER SERIA DE COLOCAR O NOME DA PRIMERA COLUNA E TIRAR TUDO ANTES

    # Abrir o arquivo txt
    with open(directory + fileName, 'r', encoding='utf-8', errors="ignore") as file:
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
        workbook = workbook.save(directory + fileName[:-4] +".XLSX")

        file.close()

    if os.path.exists(directory + fileName):
        os.remove(directory + fileName)
        print(f"O arquivo {fileName} foi deletado com sucesso.")
    else:
        print(f"O arquivo {fileName} não foi encontrado.")

        print("Planilha criada com sucesso!")
    
    return fileName[:-4] +".XLSX"


    
def JoinAndSaveNewExtract():
    
    print(fileNameExel)# = ["zse16.XLSX","Y_LAD_65000280.XLSX"]
    print("Excel para SAP")
    googleSheets = LoadFromSheets(spreadsheet_id)

    letters = string.ascii_lowercase
    t=0
    for i in fileNameExel:  #Corre pelos Arquivos exel, aplica o "remove_leading_zeros", renomeia as colunas e faz o concat
             
        df_exel = pd.read_excel(directory + i)
        if("mb51_261-262" in i):  
            googleSheets.open_sheet("1 - Download SAP - Weekly")
            worksheet = googleSheets.sheet
            range_to_clear = f'A2:C1000'

            # Execute o batch_clear para limpar o range
            worksheet.batch_clear([range_to_clear])
            df_exel = df_exel.groupby('Material').agg({'Quantity': 'sum','Amt.in Loc.Cur.':'sum'}).reset_index()
        if("mb51_7" in i):  
            t = 5
            googleSheets.open_sheet("1 - Download SAP - Weekly")
            worksheet = googleSheets.sheet
            range_to_clear = f'F2:H1000'
            worksheet.batch_clear([range_to_clear])
            
            df_exel = df_exel.groupby('Material').agg({'Quantity': 'sum','Amt.in Loc.Cur.':'sum'}).reset_index()

        if("zse16" in i):
            t = 6
            googleSheets.open_sheet("0 - SKUs info")
            worksheet = googleSheets.sheet
            range_to_clear = f'H2:L1000'

            # Execute o batch_clear para limpar o range
            worksheet.batch_clear([range_to_clear])
            df_exel= df_exel.rename(columns={"Plant": "Plnt", "MRP Controller": "MRPCn"})
            df_exel = df_exel[["Material",'Plnt','MRPCn']]

        if("Y_LAD_65000280" in i):
            t = 0
            googleSheets.open_sheet("0 - SKUs info")
            worksheet = googleSheets.sheet
            range_to_clear = f'A2:F1000'

            # Execute o batch_clear para limpar o range
            worksheet.batch_clear([range_to_clear])

            df_exel = df_exel.rename(columns={" Standard price": "Standard price", "   per":"per"})
            print(df_exel.columns)
            df_exel = df_exel[["Material","Standard price","per","BUn"]]
            
            df_exel['Standard price'] = df_exel['Standard price'].str.replace(',','')
            df_exel['Standard price'] = df_exel['Standard price'].str.strip()
            df_exel['Standard price'] = df_exel['Standard price'].astype(float)
            
            df_exel["per"] = df_exel["per"].str.replace(',','')
            df_exel["per"] = df_exel["per"].str.strip()
            df_exel["per"] = df_exel["per"].astype(float)
            
            df_exel["STD Cost"] = df_exel["Standard price"] / df_exel["per"]
           # df_exel = df_exel.loc[:, ~df_exel.columns.str.contains('^Unnamed')]
            
  
        df_exel = df_exel.sort_values(by='Material')
        df_exel = pd.DataFrame(df_exel)
        googleSheets.df_to_sheet(df_exel, sheet=worksheet,index=False, start = letters[0+t]+ '2') 
        print(i + " extraiu na " +letters[0+t]+ "2") 
        


def AutoRun():
    if (autoRun.get() == 1):
        print("Automatico!")
        c1.config(state=tk.DISABLED)
        SAP_Extraction()
        time.sleep(3)
        JoinAndSaveNewExtract()
        c1.config(state= tk.NORMAL)
        
    
def GUIparaVariavel():
    
    global system_SAP, directory, spreadsheet_id  # Suponho que essas variáveis globais já existam

    # Coletar os dados dos campos da interface gráfica (Tkinter)
    system_SAP = campo_sistema_sap.get()
    directory = campo_caminho_diretorio.get()
    spreadsheet_id = campo_spreadsheet_id.get()

    
# Função para salvar as informações em um arquivo JSON
def salvar_informacoes():

    GUIparaVariavel()
    # Criar um dicionário com os dados
    dados = {
        "Sistema do SAP": campo_sistema_sap.get(),
        "Caminho do Diretório": campo_caminho_diretorio.get(),
        "SpreadSheet ID": campo_spreadsheet_id.get(),
        "AutoRun": autoRun.get()  
    }
    
        # Verificar se o arquivo JSON já existe
    if not os.path.exists("informacoes.json"):
        # Se não existir, criar o arquivo e salvar os dados
        with open("informacoes.json", "w") as arquivo:
            json.dump(dados, arquivo, indent=4)
        print("Arquivo 'informacoes.json' criado e informações salvas com sucesso!")
    else:
        # Se o arquivo já existir, apenas atualizar os dados
        with open("informacoes.json", "w") as arquivo:
            json.dump(dados, arquivo, indent=4)
        print("Informações atualizadas e salvas no arquivo 'informacoes.json'.")

    # Salvar os dados em um arquivo JSON
    with open("informacoes.json", "w") as arquivo:
        json.dump(dados, arquivo, indent=4)

    print("Informações salvas com sucesso!")

# Exemplo de como carregar as informações
def Carregar_informacoes():
    if os.path.exists("informacoes.json"):
        with open("informacoes.json", "r") as arquivo:
            dados = json.load(arquivo)

            campo_sistema_sap.delete(0, tk.END)
            campo_sistema_sap.insert(0, dados["Sistema do SAP"])

            campo_caminho_diretorio.delete(0, tk.END)
            campo_caminho_diretorio.insert(0, dados["Caminho do Diretório"])

            campo_spreadsheet_id.delete(0, tk.END)
            campo_spreadsheet_id.insert(0, dados["SpreadSheet ID"])
            
            autoRun.set(dados["AutoRun"])
                               
            GUIparaVariavel()
    else:
        print("Arquivo de informações não encontrado.")
        
        
if __name__ == '__main__':
            
    window = tk.Tk()
    window.geometry("500x400")
    window.title("Configurações")

    # Frame para conter todos os widgets
    frame = tk.Frame(window)
    
    # Estilo para os títulos
    titulo_fonte = font.Font(family="Helvetica", size=14, weight="bold")

    # Criar os campos de entrada
    tk.Label(window, text="Extração do SAP v1", font=titulo_fonte).pack(pady=10)

    tk.Label(window, text="Sistema do SAP:", font=("Helvetica", 10)).pack()
    campo_sistema_sap = Entry(window, bg="#E8E8E8", font=("Helvetica", 10))
    campo_sistema_sap.pack(padx=10, pady=5)
    
    tk.Label(window, text="Caminho do Diretório:", font=("Helvetica", 10)).pack()
    campo_caminho_diretorio = Entry(window, bg="#E8E8E8", font=("Helvetica", 10), width=30)
    campo_caminho_diretorio.pack(padx=10, pady=5)

    tk.Label(window, text="SpreadSheet ID:", font=("Helvetica", 10)).pack()
    campo_spreadsheet_id = Entry(window, bg="#E8E8E8", font=("Helvetica", 10), width=50)
    campo_spreadsheet_id.pack(padx=10, pady=5)
    
    autoRun = tk.IntVar()
    c1 = tk.Checkbutton(window, text='Automatic',variable=autoRun, onvalue=1, offvalue=0,)
    c1.pack()   
    # Carregar informações se existirem
    Carregar_informacoes()

    # Botão para salvar as informações
    botao_salvar = Button(window, text="Salvar Informações", command=salvar_informacoes, bg="#4CAF50", fg="white", font=("Helvetica", 10, "bold")).pack(pady=10)
    # Botões de funcionalidade
    frame_botoes = tk.Frame(window)
    frame_botoes.pack(pady=15)
    Button(frame_botoes, text="Login SAP", command=SAP_Extraction, bg="#008CBA", fg="white", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=5)
    Button(frame_botoes, text="Excel para Sheets", command=JoinAndSaveNewExtract, bg="#008CBA", fg="white", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=5)

    AutoRun()
    window.mainloop()
    
