#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import pandas as pd
import time
import json
import smtplib
import urllib.parse
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders


# In[ ]:


path = '' #Path to download file
end_file = ' - Servidor de Relatórios do Power BI.pdf'
shrpt_url = 'https://xxx.sharepoint.com'
shrpt_folder = 'Pasta xxx'
shrpt_site = 'site/xxx'
pbi_user = os.getenv("email_pbi")
pbi_pass = os.getenv("password_pbi")


# In[ ]:


def authentication(pbi_user, pbi_pass, shrpt_url, shrpt_site):
    ctx_auth = AuthenticationContext(shrpt_url)
    ctx_auth.acquire_token_for_user(pbi_user, pbi_pass)
    ctx = ClientContext(f'{shrpt_url}/{shrpt_site}/', ctx_auth)
    return ctx


# In[ ]:


def list_files_shrpt(auth) -> list:
    try:
        folder_url = urllib.parse.quote(shrpt_folder)
        context_object = auth
        folder = context_object.web.get_folder_by_server_relative_url(folder_url)
        folders_names = []
        sub_folders = folder.files
        context_object.load(sub_folders)
        context_object.execute_query()

        for sub_folder in sub_folders:
            f_name = sub_folder.properties['Name']
            folders_names.append(f_name)         
        return folders_names
    except Exception as e:
        print(e)


# In[ ]:


def download_file_shrpt(file_name,path):
    try:
        file_url = f'/{ shrpt_site}/{ shrpt_folder}/{file_name}'
        download_path = os.path.join(path, os.path.basename(file_url))
        with open(download_path, "wb") as local_file:
            file = authentication(pbi_user, pbi_pass, shrpt_url, shrpt_site).web.get_file_by_server_relative_path(file_url).download(local_file).execute_query()
        print("[Ok] file has been downloaded into: {0}".format(download_path))
        return download_path
    except Exception as e:
        raise ValueError(e)


# In[ ]:


def download_report(report_list,pbi_user,pbi_pass):
    try:
            #Configures to print and save the report without having to click save and enter the name
        chrome_options = webdriver.ChromeOptions()
        settings = {
              "recentDestinations": [{
                   "id": "Save as PDF",
                   "origin": "local",
                   "account": "",
               }],
               "selectedDestinationId": "Save as PDF",
               "version": 2,
               "isCssBackgroundEnabled": True,
               "isHeaderFooterEnabled": True
        }
        prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument('--kiosk-printing')
        s = Service('.../chromedriver') #Path to ChromeDriver
    
           #Open the browser
        browser = webdriver.Chrome(options=chrome_options, service=s)
        print('Chrome aberto...')
        time.sleep(5)

            #I entered the username and password directly in the url to access the report page and then go straight to the home page
        browser.get(f'https://{pbi_user}:{pbi_pass}@reportserver.com.br/reports/browse/Repositorio')
        browser.get('https://relatorios.xxx.com.br/reports/browse/Reposit')
        print('Iniciar Login... \nInserindo email e senha!')
        time.sleep(5)
    
            #Get All Report in list
        for report in report_list:
            browser.get(report)
            time.sleep(10)

                #Click on File
            browser.find_element(By.XPATH,'//').click()
            time.sleep(10)
   
                #Click on Print
            browser.find_element(By.XPATH,'//').click()
            time.sleep(10)
            print('Relatório Salvo... \nEnviando email!')
    
            #Report save, close the browser.
        browser.close()
    except Exception as e:
        raise ValueError(e)


# In[ ]:


def send_email(file, receiver,name_report):
    try:
        body_email = f'''
        <p>Olá!</p>

        <p>Segue, em anexo, seu report {name_report}.</p>

        <p>Caso tenha qualquer dúvida, estamos à disposição!</p>

        <p>Abraços,<br>
        Equipe BI.</p>
        <p>***Email enviado automaticamente!***</p>
        '''

        msg = MIMEMultipart()
        msg['Subject'] = f'{name_report} - (Report Server)'
        msg['From'] = '' #Email to send reports

        text = MIMEText(body_email, 'html')
        msg.attach(text)

        smtp_server = "smtp.gmail.com"
        port = 587  # For starttls

        # Try to log in to server and send email
        server = smtplib.SMTP(smtp_server, port)
        server.ehlo()  # check connection
        server.starttls()  # Secure the connection
        server.ehlo()  # check connection
        server.login(msg['From'], "xxxxx") #Insert Password App

        # Attachment
        attach = MIMEBase('application', "octet-stream")
        attach.set_payload(file.read())
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition', 'attachment; filename="Report.pdf"')
        msg.attach(attach)

        # Send email here
        server.sendmail(msg['From'], receiver, msg.as_string())

        server.quit()
    except Exception as e:
        raise ValueError(e)


# In[ ]:


def start_sends():
    try:
        print(datetime.now(), 'Autenticando no Sharepoint...')
        auth = authentication(pbi_user, pbi_pass, shrpt_url, shrpt_site)
        print(datetime.now(), 'Listando os arquivos no diretório do Sharepoint...')
        lst_filesShrpt = list_files_shrpt(auth)
        print(datetime.now(), 'Realizando download do Sharepoint...')
        for file_name in lst_filesShrpt:
            print(datetime.now(), f'Download do arquivo: {file_name} do Sharepoint..')
            download_path = download_file_shrpt(file_name,path)
            urls = pd.read_excel(download_path,sheet_name='URL')
            download_report(urls['URL'],pbi_user,pbi_pass)
            for name_report in urls['Report']:
                pdf = path+name_report+end_file
                email = pd.read_excel(download_path,sheet_name=name_report)
                if os.path.exists(pdf):
                    file = open(pdf,'rb')
                    send_email(file,email['Email'],name_report)
                    print(f'Email relatório {name_report} enviado!')
                    file.close()
                    os.remove(pdf)
                else:
                    print('Arquivo não existe!')
    except Exception as e:
        raise ValueError(e)


# In[ ]:


print(datetime.now(),f'**** INÍCIO ****')
start_sends()
print(datetime.now(),f'**** FIM ****')  


# In[ ]:




