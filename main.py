import time, schedule, smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from passwords import emailPassword

def sendEmail(managerNumber, password = emailPassword):
    #definindo gerente
    managerNumber = managerNumber

    managerEmail = ['managerEmail1', 'managerEmail2', 'managerEmail3']

    # Configurações do servidor SMTP e credenciais
    smtp_server = 'smtp.gmail.com'
    port = 587

    # Cpnteudo auth https://youtu.be/jHP-9fXQGwo?si=6AWDSTiVXFYKxEWu
    emailFrom = 'yourEmail'
    emailPassword = password
    emailTo = managerEmail[managerNumber-1]

    
    #caminho dos arquivos de anexo
    archiveName1 = f'gerente {managerNumber} - 1'
    archiveName2 = f'gerente {managerNumber} - 2'
    archiveName3 = f'gerente {managerNumber} - 3'
    archivePath1 = f'C:\\Users\\Iago Lima\\Desktop\\rpa-email\\gerente {managerNumber}\\{archiveName1}.pdf'
    archivePath2 = f'C:\\Users\\Iago Lima\\Desktop\\rpa-email\\gerente {managerNumber}\\{archiveName2}.pdf'
    archivePath3 = f'C:\\Users\\Iago Lima\\Desktop\\rpa-email\\gerente {managerNumber}\\{archiveName3}.pdf'

    #abrindo o arquivo em modo leitura e como binary
    pdf1 = open(archivePath1, 'rb')
    pdf2 = open(archivePath2, 'rb')
    pdf3 = open(archivePath3, 'rb')

    #lendo como binary e transformando em base 64
    #transformando pdf 1
    attachment1 = MIMEBase('application', 'octet-stream')
    attachment1.set_payload(pdf1.read())
    encoders.encode_base64(attachment1)
    #transformando pdf 2
    attachment2 = MIMEBase('application', 'octet-stream')
    attachment2.set_payload(pdf2.read())
    encoders.encode_base64(attachment2)
    #transformando pdf 3
    attachment3 = MIMEBase('application', 'octet-stream')
    attachment3.set_payload(pdf3.read())
    encoders.encode_base64(attachment3)

    #cabecalho dos anexos 1, 2 e 3
    attachment1.add_header('Content-Disposition', f'attachment; filename= {archiveName1}.pdf')
    attachment2.add_header('Content-Disposition', f'attachment; filename= {archiveName2}.pdf')    
    attachment3.add_header('Content-Disposition', f'attachment; filename= {archiveName3}.pdf')
    
    #fechando arquivos
    pdf1.close()
    pdf2.close()
    pdf3.close()
    
    # Configuração do email
    emailSubject = f'Olá gerente {managerNumber}'
    emailBody = f'<div style="font-family: Arial, sans-serif;background-color:#f4f7fc; margin: 0; padding: 0;"><div style="max-width: 600px; margin: 20px auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; background-color: #fff; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);"><h1 style="color: #007BFF; border-bottom: 2px solid #007BFF; padding-bottom: 10px; margin-bottom: 20px;">Relatório Diário para Conferência de PDFs</h1><div style="margin-top: 15px;"><p style="color: #333; margin: 0; padding: 5px 0;">Segue anexos para a conferência por parte do gerente:</p><p style="color: #333; margin: 0; padding: 5px 0;">- Relatório 1: {archiveName1}</p><p style="color: #333; margin: 0; padding: 5px 0;">- Relatório 2: {archiveName2}</p><p style="color: #333; margin: 0; padding: 5px 0;">- Relatório 3: {archiveName3}</p></div></div></div>'


    message = MIMEMultipart()
    message['From'] = emailFrom
    message['To'] = emailTo
    message['Subject'] = emailSubject
    message.attach(MIMEText(emailBody, 'html'))

    #anexando arquivos ao email
    message.attach(attachment1)
    message.attach(attachment2)
    message.attach(attachment3)
    
    
    # Conexão com o servidor SMTP
    server = smtplib.SMTP(smtp_server, port)
    server.starttls()
    server.login(emailFrom, emailPassword)

    # Envio do email
    server.sendmail(emailFrom, emailTo, message.as_string())

    # Encerramento da conexão com o servidor
    server.quit()

    print("Email enviado com sucesso!")

def dailySending():
    for i in range(1, 4):
        sendEmail(i)
    
schedule.every().day.at("23:45").do(dailySending)

while True:
        schedule.run_pending()
        time.sleep(1)
