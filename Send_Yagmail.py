#-------------------------------------
#  SEND email
# https://stackoverflow.com/questions/10147455/how-to-send-an-email-with-gmail-as-provider-using-python/27515833#27515833
# Tested OK
# https://code.tutsplus.com/pt/tutorials/sending-emails-in-python-with-smtp--cms-29975
# https://realpython.com/python-send-email/
# https://www.programcreek.com/python/example/103416/email.mime.image.MIMEImage
# https://stackoverflow.com/questions/34008313/error-in-customizing-receivers-name-in-automated-email-script-in-python https://blog.mailtrap.io/yagmail-tutorial/
# To create .ics eventos: https://ical.marudot.com/

# pip install mimelib
# pip install yagmail

import smtplib
import os
import numpy as np
import pandas as pd
import yagmail
import random
import time
from time import gmtime, strftime
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

#--------- EMAIL SYSTEM
# create message object instance
username = 'manuel.robalinho@gmail.com'
# a password é obtiga na configuração de segurança em 2 passos do google
# ao definir password para acesso de APPs
password = 'fxsaglakxrxztqfc'
fromaddr = 'manuel.robalinho@gmail.com' 

yag = yagmail.SMTP(fromaddr, password)

#------------------------------------------------------------
# Send Email function
def send_yagmail(fromaddr, toaddr, email_subject, body, path, file1, file2, file3):
    print('---->> Sending mail to :', toaddr, '  Subject:', email_subject)

    # Send email
    if file1 is not '' and file2 is not '' and file3 is not '':
        yag.send(to=toaddr, subject=email_subject, contents=body, attachments=[path+file1, path+file2, path+file3])
        print('---->> Sent the mail and 3 attachments to :', toaddr)
    else: 
        if file1 is not '' and file2 is not '':
            yag.send(to=toaddr, subject=email_subject, contents=body, attachments=[path+file1, path+file2])
            print('---->> Sent the mail and 2 attachments to :', toaddr)
        else:
            if file1 is not '':
                yag.send(to=toaddr, subject=email_subject, contents=body, attachments=[path+file1])
                print('---->> Sent the mail and 1 attachments to :', toaddr)
            else:
                yag.send(to=toaddr, subject=email_subject, contents=body)
                print('---->> Sent the mail without attachments to :', toaddr)
    
    #time.sleep(random.randrange(2,5))
    #print('Sent the mail to :', toaddr)
    return
#-------------------------------------------
def create_msg_email(nome, xdoc, xdata, xemail, xvalor, xfile_pdf, xfile_jpg, xfile_event):
    
    path = 'files_to_attach/'
    email_subject = 'SAP Inside Track Fortaleza 2020'
    link = 'https://blogs.sap.com/2020/05/31/sap-inside-track-fortaleza-1a-edicao/'
   #------------------------------- 
    toaddr  = xemail
    print('2.Create msg :',nome, xdoc, xdata, xvalor)
    
    # Identificação do doc
    doc   = xdoc
    data_doc  = xdata
    valor = xvalor
    toname = nome
   
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

    email_subject = email_subject + ' , em '+str(data_doc)

   # Create Body email

    #template = 'Hello {name}, it is me'           # yagmail automatically makes this HTML
    template =  "\r\n".join([
        '<h2>Olá {name}, seja bem vindo !</h2> <br>',
        "Informamos que o evento <b>"+str(doc)+"</b> vai ocorrer na data: {data}",
        "Fique por dentro das palestras.",  
        "<a href='" +link+ "'>" +"Link para o evento" + "</a>",
        "",  
        "Cumprimentos,",
        "<b>Manuel Robalinho.</b>" ,
        "",  
        "Sent at:"+strftime("%Y-%m-%d %H:%M:%S", gmtime()) 
        ]) 

    body = template.format(name=toname, data=data_doc)
    
    # Call function to send email
    print('3.Sent the mail to :', xemail)
    
    send_yagmail(fromaddr, toaddr, email_subject, body, path, xfile_pdf, xfile_jpg, xfile_event )
    
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

