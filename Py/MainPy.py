#!/usr/bin/env python
# coding: utf-8

# In[1]:


#Import vari

import smtplib, ssl
import os as os
import openpyxl
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# In[2]:


#Credenziali e percorsi

# Definisci le tue credenziali di accesso al tuo account email
email = 'commerciale@4digits.it' 
password = 'kH10J3nLM!bd'
nMailDaInviare = 20
oggettoEmail_path = '../Oggetto_Email.txt'

#Definisci dati del file excel da cui prendere le mail
nomeExcelDestinatari = '../Excel/Dentist In Italy .xlsx'
sheet_name='Sheet1'
colonnaEmail = 'Business Email_1'

#Definisci dati del file excel dove scrivere i dati dei dentisti già "usati"
nomeExcelInviati = '../Excel/Inviati.xlsx'

#Contenuto mail
testo = '../Testo_Email.txt'

#Codice footer html
html = 'footer.html'    


# In[3]:


#Pulisce il file con le email già usate e le salva su un nuovo file

def PulizziaFile():
    excelGrezzo = pd.read_excel(nomeExcelDestinatari, sheet_name) #Test
    excelDestinatari = excelGrezzo.head(nMailDaInviare) #prende i primi 20 a cui mandare oggi la mail
    excelNuovo = excelGrezzo.iloc[20:] #elimina i primi 20 dal grezzo
    excelNuovo.to_excel(nomeExcelDestinatari, sheet_name, index=False) #sovrascrive tutto su iviati.excel
    excelInviati = pd.read_excel(nomeExcelInviati, sheet_name) #legge gli inviati in precedenza
    df_updated = pd.concat([excelInviati, excelDestinatari], ignore_index=True) #unisce quelli attuali con quelli inviati in precedenza
    df_updated.to_excel(nomeExcelInviati, sheet_name, index=False) #sovrascrive tutto su iviati.excel




# In[4]:


# Leggi l'oggetto del messaggio dal file di testo
with open(oggettoEmail_path, 'r') as file:
    oggettoEmail = file.read()


# In[5]:


# Leggi il testo del messaggio dal file di testo
with open(testo, 'r') as file:
    testo_messaggio = file.read()


# In[6]:


#Legge il codice html per il footer

with open(html, 'r') as file:
    codiceHtml = file.read()


# In[7]:


#Funzione creazione messaggio

def creaMessaggio():
    # Crea il messaggio
    messaggio = MIMEMultipart()
    messaggio['From'] = "4Digits<"+email+">"
    messaggio['Subject'] = oggettoEmail
    # Aggiungi il testo del messaggio al corpo dell'email

    #Aggiungi footer al messaggio


    messaggio.attach(MIMEText(testo_messaggio, 'plain'))
    
    messaggio.attach(MIMEText(codiceHtml, "html"))


    
    

    return messaggio


# In[8]:


# Invia l'email a ciascun destinatario
def invioMail(destinatari):

    for destinatario in destinatari:
        
        messaggio = creaMessaggio()
        
        try:
            messaggio['To'] = destinatario
            #server = smtplib.SMTP('smtp.4digits.it', 25) # Imposta il server SMTP del tuo provider email
            
            server = smtplib.SMTP_SSL('smtps.aruba.it', 465) # Imposta il server SMTP del tuo provider email
            
            #server = smtplib.SMTP('smtps.aruba.it', 465) # Imposta il server SMTP del tuo provider email
            #server.starttls()
            server.login(email, password)
            testo_email = messaggio.as_string()
            server.sendmail(email, destinatario, testo_email)
            server.quit()
        except Exception as e:
        # Print any error messages to stdout
                print(e)
            


# In[9]:


# Leggi i primi 20 destinatari dal file Excel
destinatari = pd.read_excel(nomeExcelDestinatari, sheet_name)
destinatari = destinatari[colonnaEmail].head(nMailDaInviare).tolist()
destinatari.append("andreettonicolo@gmail.com")
print(destinatari)


invioMail(destinatari) #Richiama anche creaMessaggio

PulizziaFile()

print('Le email sono state inviate correttamente')
os.system("pause")



# In[ ]:





# In[ ]:




