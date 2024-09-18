from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import os

# Configurações Globais
fileNameExel = ["Y_LAD_65000872.xlsx"]

home_directory = os.path.expanduser("~")

local_path = home_directory + "\\Desktop\\"

spreadsheet_id = '1YV_GeWsY8H6GemikNZQCxZ9TbkUdCrQ1bq0XBRy2nnY'
folder_id = '1vBqoEyHAriKeJcPzsOXbn8XjJiUXLz2m'
control_spreadsheet_id="18VCyPwR1G5AhXBCBKEJ0WCnlq5gfqacJAMKMuWgQgP8"

system_SAP = "LAP"
env = "dev"


service_account_info = {       

      }

# Pega primeira e ultima data do mes atual
dateToday = datetime.now()

startDate = dateToday.replace(day=1)
endDate = startDate - timedelta(days=1)
endDate += relativedelta(months=1)

startDateFull = startDate.strftime('%m%d%Y')
endDateFull  = endDate.strftime('%m%d%Y')

startDateMonth = startDate.strftime('%m%d%Y')
endDateMonth = endDate.strftime('%m%d%Y')

startDateMonth = startDate.strftime('%m')
endDateMonth = endDate.strftime('%m')

startDateYear =  startDate.strftime('%Y')
endDateYear =  startDate.strftime('%Y')

