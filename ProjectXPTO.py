import os
import shutil
import glob
import PyPDF2
import smtplib
import datetime
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# Coletar data e hora
data_e_hora_atuais = datetime.now()
data_e_hora_em_texto = data_e_hora_atuais.strftime('\n\nDia: %d/%m/%Y \nHorario da execuçao: %H:%M')
data_log = data_e_hora_atuais.strftime('%d/%m/%Y %H:%M')
hora = data_e_hora_atuais.strftime('%H:%M')


# Diretorio do documento do log
notas_sem_fatura = ('D:\\INTEGRACAO\\CheckFatura\\notas sem fatura.txt')


# Verificar se o arquivo de "registro de notas sem fatura" já existe, caso não exista ele cria um arquivo novo

if os.path.exists(notas_sem_fatura):
    fh = open ('notas sem fatura.txt','a') # abrir arquivo
else:
    try:
        fh = open ('notas sem fatura.txt','w') #criar arquivo
    except:
        print('Falha ao tentar criar arquivo')


# Diretorio onde ficam os PDFs para leitura

caminho_dos_pdfs = '~\\PDFTESTE\\'

for i in glob.glob(caminho_dos_pdfs+'*.PDF'):
    i = i.replace('~\\PDFTESTE\\','')
    try:
        file = open('~\\PDFTESTE\\'+i,'rb')
    except:
        print('')
    # Ler o documento .PDF (ele lê somente a 1 pg do PDF)
    pdfreader=PyPDF2.PdfFileReader(file, strict=False)
    pageobj=pdfreader.getPage(0)
    texto_pdf = pageobj.extractText()
    palavra_chave = 'VENDA'
    print('.',end=' ')

    # Verificar se o arquivo ja foi lido
    nf_sem_fatura = open('notas sem fatura.txt','r')
    notas = nf_sem_fatura.read()
    if i in notas:
        print('+',end=' ')
        continue

    # Condicionais onde a nota pode não ter fatura
    if palavra_chave in texto_pdf:
        if 'DEVOLUCAO DE VENDA' in texto_pdf:
            continue
        if 'ANULACAO VALOR RELATIVO' in texto_pdf:
            continue
        if 'REMESSA DE VASILHAME OU SACARIA' in texto_pdf:
            continue
        if 'REMESSA  MERCADORIA P/ CONTA E ORDEM TERC. VENDA ORDEM' in texto_pdf:
            continue
        if 'REMESSA PARA INDUSTRIALIZACAO POR ENCOMENDA' in texto_pdf:
            continue

        # Caso ele realmente for uma venda ele vai seguir o processo
        if not 'FATURA' in texto_pdf:
            # Escrever a nota no txt de notas_enviadas
            fh.write(i+'   - '+data_log+'\n')

            # Enviar o e-mail informando que a venda não possui fatura
            remetente = 'ERP@'
            destinatarios = 'EMAILS@'
            msg = MIMEMultipart()
            msg['From'] = remetente
            msg['To'] = str(destinatarios)
            msg['Subject'] = 'VENDA SEM FATURA'
            i = i.replace('.PDF','')

            body = 'A nota '+i+' não possui fatura, favor verificar. \n \n \n \n \n Horario de execução: '+hora+'\n Palavra chave: '+palavra_chave
            msg.attach(MIMEText(body, 'plain'))

            # Configuração so servidor SMTP
            s = smtplib.SMTP('smtp.office365.com', 587)
            s.starttls()
            s.login(remetente, "senha")
            text = msg.as_string()
            s.sendmail(remetente, destinatarios, text)
            s.quit()

            print('\nARQUIVO '+i+' NÃO POSSUI FATURA, EMAIL ENVIADO')

        else:
            continue

# Fechar documento de texto para salvar alterações
fh.close()