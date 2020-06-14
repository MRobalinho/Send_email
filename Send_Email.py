#-------------------------------------
#  SEND email
# https://stackoverflow.com/questions/10147455/how-to-send-an-email-with-gmail-as-provider-using-python/27515833#27515833
# Tested OK
# https://code.tutsplus.com/pt/tutorials/sending-emails-in-python-with-smtp--cms-29975
# https://realpython.com/python-send-email/
# https://www.programcreek.com/python/example/103416/email.mime.image.MIMEImage
# To create .ics eventos: https://ical.marudot.com/

# pip install mimelib

import smtplib
import os
import numpy as np
from time import gmtime, strftime 
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
import pandas as pd
#------------------------------------------------------
# File Excel
xls_file = 'data_to_send.xlsx'
print('-----  ENVIO DE EMAILS ------')
print('Lendo arquivo :', xls_file)

df = pd.read_excel('data_to_send.xlsx', sheet_name='Sheet1')

# print whole sheet data
df

df.columns = ['cliente', 'nome', 'zz', 'doc', 'item', 'referencia', 'vencimento', 'montante', 'email','file_pdf','file_jpg','file_event']
# Only rows with email
df2 = df.dropna(subset=['email']).copy()
df2.head()

# for column NaN values
df2['file_pdf'].replace(np.nan, '', inplace=True)
df2['file_jpg'].replace(np.nan, '', inplace=True)
df2['file_event'].replace(np.nan, '', inplace=True)

# create message object instance
msg = MIMEMultipart()


#------------------------------------------------------------
# Send Email function
def send_email(fromaddr, toaddrs, msg):
    username = 'manuel.robalinho@gmail.com'
    # a password é obtiga na configuração de segurança em 2 passos do google
    # ao definir password para acesso de APPs
    password = 'fxsaglakxrxztqfc'
    #....
    server = smtplib.SMTP_SSL("smtp.gmail.com",465)
    server.ehlo()
    server.login(username, password)
    server.sendmail(fromaddr, toaddrs, msg)
    server.close()
    print('4.Successfully sent the mail to :',toaddrs)

    return
#-------------------------------------------
def create_msg_email(nome, xdoc, xdata, xemail, xvalor, xfile_pdf, xfile_jpg, xfile_event):
    msg = MIMEMultipart()
    
    path = 'files_to_attach/'
    
    fromaddr = 'manuel.robalinho@gmail.com'
    toaddrs  = xemail
    print('2.Create msg :',nome, xdoc, xdata, xvalor)
    
    # Identificação do doc
    doc   = xdoc
    data  = xdata
    valor = xvalor
   
    # Image File
    if xfile_jpg is not '':
        image_file = path + xfile_jpg
        img_data = open(image_file, 'rb').read()
    
    # Pdf File
    file_pdf = ''
    if xfile_pdf is not '':
        file_pdf = path + xfile_pdf
    
    # Calendar Event file
    file_event = ''
    if xfile_event is not '':
        file_event = path + xfile_event    
    #
    password = "your_password"
    msg['From'] = fromaddr
    msg['To']   = toaddrs
    msg['Subject'] = " Vencimento do documento:"+str(doc)+" em "+data+ "  Valor:"+str(valor)

    # Mensagem HTML
    message = "\r\n".join([
      "From: "+fromaddr+"<br>",
      "To: "+toaddrs+"<br>",
      "<br>",
      "<h3>Subject: "+msg['Subject']+"</h3>",
      "<br>",
      "Exmo Sr: <b>"+ nome +'</b>'+ "<br>",
      "Informamos que o documento <i>"+str(doc)+"</i> no valor de "+str(valor)+" ,vence em "+data+" .",
      "Agradecemos o pagamento atraves da conta bancaria ABCD.",
      "<br>",
      "<br>",
      "<br>",
      "Cumprimentos,<br>",
      "Grupo Aco Cearense.<br>" ,
      "<br>",  
      "Sent at:"+strftime("%Y-%m-%d %H:%M:%S", gmtime())  
      ])

    # add in the message body
    #text = MIMEText(message,'plain')
    text = MIMEText(message,'html')
    msg.attach(text)

    # attach image to message body
    if xfile_jpg is not '':
        print('......Tem jpg :', image_file)
        image = MIMEImage(img_data, name=os.path.basename(image_file))
        msg.attach(image)

    # attach PDF to email - OK
    if xfile_pdf is not '': 
        print('......Tem pdf :', file_pdf)
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(file_pdf, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % file_pdf)
        msg.attach(part)

    # attach Event Calendar - OK
    if xfile_event is not '': 
        print('......Tem evento :', file_event)
        part = MIMEBase('text', 'calendar',method='REQUEST',name=file_event) #method ='REQUEST' only provide me the possibility of adding 1 event, not 5
        part.set_payload(open(file_event, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % xfile_event) 
        msg.attach(part)
		
    # Call function to send email
    print('3.Sent the mail to :', xemail)
    
    send_email(fromaddr, xemail, msg.as_string() )
    return
#------------------------------------------------------------------------
def executa_row(nome, doc, data, email, valor, file_pdf, file_jpg, file_event):
    print('1.Exec. Row: ', nome,  email)

    create_msg_email(nome, doc, data, email, valor, file_pdf, file_jpg, file_event)
    return

# teste
# executa_row('manuel SIlva', 'doc123', '13-6-2020', 'manuel.robalinho@hotmail.com', 20000.55, 'file_pdf', 'file_jpg', 'file_event'   )
# Read each line from Dataframe
for index, row in df2.iterrows():
    print('...') 
    print('---> ', index+1)
    executa_row(row['nome'], row['doc'], row['vencimento'], row['email'], row['montante'], row['file_pdf'], row['file_jpg'] , row['file_event'])
    
print('--- Finished ----')

