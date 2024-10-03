from datetime import datetime, timedelta  # Importa módulos para manipular datas.
from dateutil.relativedelta import relativedelta  # Importa um módulo para manipulação avançada de datas.
import os  # Importa módulo para interações com o sistema operacional.

# Configurações Globais
fileNameExel =["md04_B0750351M.xlsx", "md04_B07503020001.xlsx","md04_B07505850000.xlsx"]

                                          #Isso é para o Serve e Local
home_directory = os.path.expanduser("~")  # Obtém o diretório home do usuário atual.
local_path = home_directory + "\\Desktop\\"  # Define o caminho local como a área de trabalho do usuário.

control_spreadsheet_id = "18VCyPwR1G5AhXBCBKEJ0WCnlq5gfqacJAMKMuWgQgP8"#Esse é o Sheets de Log de Controle

folder_id = "1WRWHxyfLiMIBmlVmFxDANk1CKmkJFgUV"
spreadsheet_id= "1E4UGPw7lV73b8GPxLfPNxNtS9hGb_lbPPNCLSjQ80Qs"
system_SAP = "LAP"  # Sistema SAP utilizado.


project_name  ="Alex_nmd04"
env = "prod"  # Ambiente de desenvolvimento.#Isso serve para a Sheets de controle

#GOOGLE API
service_account_info ={  

      }  # Informações da conta de serviço (atualmente vazio).

# Pega a primeira e a última data do mês atual.
dateToday =datetime(2024,9,20) # # Obtém a data e hora atual.
# Define o primeiro dia do mês atual.
startDate = dateToday.replace(day=1)

# Calcula o último dia do mês atual.
# Primeiro, define-o como o último dia do mês anterior e depois ajusta para o mês corrente.
endDate = startDate - timedelta(days=1)
endDate += relativedelta(months=1)

dateToday2 =datetime.now()  
startDate2 = dateToday2.replace(day=1)

# Calcula o último dia do mês atual.
# Primeiro, define-o como o último dia do mês anterior e depois ajusta para o mês corrente.
endDate2 = startDate2 - timedelta(days=1)
endDate2 += relativedelta(months=1)
# Formata as datas como strings no formato mmddyyyy.
startDateFull = startDate.strftime('%m%d%Y')
endDateFull = endDate.strftime('%m%d%Y')

startDateFull2 = startDate2.strftime('%m%d%Y')
endDateFull2 = endDate2.strftime('%m%d%Y')

# Aqui, as variáveis estão definidas duas vezes de forma desnecessária.
# Primeira definição: Formato mmddyyyy.
startDateMonth = startDate.strftime('%m%d%Y')
endDateMonth = endDate.strftime('%m%d%Y')

# Segunda definição: Somente o mês.
startDateMonth = startDate.strftime('%m')
endDateMonth = endDate.strftime('%m')

# Obtém o ano.
startDateYear = startDate.strftime('%Y')
endDateYear = endDate.strftime('%Y')
