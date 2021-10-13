import shutil
import os
from datetime import datetime

data_e_hora_atuais = datetime.now()
log_name = data_e_hora_atuais.strftime('%Y%m%d.txt')

os.rename('notas sem fatura.txt',log_name)

diretorio = os.getcwd()
diretorio = diretorio.replace('\\','\\\\')

diretorio_log = diretorio+('\\\Log')
print(diretorio_log)

shutil.move(log_name,diretorio_log)

