# %%
import requests
import pandas as pd
import time
import sqlite3
from bs4 import BeautifulSoup
from selenium import webdriver
import os
import re
from datetime import datetime
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# % definicja funkcji %
def rachunek(komornik_konto):
    # sprawdzenie długości numeru rachunku
    bank = komornik_konto[2:10]
    komornik_konto = re.sub("[^0-9]+", "", komornik_konto)
    # obsługa jednego wskazanego rachunku bankowego
    if len(komornik_konto) == 26:
        if int(komornik_konto[2:] + "2521" + komornik_konto[0:2]) % 97 == 1:
            f_konto = "1"
        else:
            f_konto = "0"
        #print("Dodano rachunek 26:", komornik_konto)
        time.sleep(1)
    # obsługa dwóch wskazanych rachunków bankowych
    if len(komornik_konto)>26:
        komornik_konto = komornik_konto[:26]
        if int(komornik_konto[2:] + "2521" + komornik_konto[0:2]) % 97 == 1:
            f_konto = "1"
        else:
            f_konto = "0"
        #print("Dodano rachunek ponad 26:", komornik_konto)
        time.sleep(1)
    # obsługa braku rachunku bankowego
    if len(komornik_konto)==0:
        komornik_konto = "00000000000000000000000000"
        f_konto = "0"
        #print("Dodano rachunek:", komornik_konto)
        time.sleep(1)
    return re.sub("[^0-9]+", "", komornik_konto)

# % główny program %
dataDanych = datetime.today().strftime('%Y%m%d')

page = requests.get('https://www.komornik.pl/?page_id=195')
soup = BeautifulSoup(page.content, "html.parser")

lista_izby = soup.find('div', class_="lista")

# tworzę pustą listę
dict_izby = []
dict_komornicy = []
dict_sady = []
logDrill = []

for a in lista_izby.find_all('a', href=True):
    izba_id = (a['href'][9:])
    izba_name = (a.get_text())
    izby = {
        'IzbaId': izba_id,
        'IzbaNazwa':izba_name
            }
    # wrzucenie słownika do pustej listy
    dict_izby.append(izby)

df_dict_izby = pd.DataFrame(dict_izby)
izba_id_list = df_dict_izby['IzbaId'].tolist()

for i in izba_id_list:
    # uruchamiam przeglądarkę
    driver = webdriver.Safari()
    driver.get('https://www.komornik.pl/?page_id='+str(i))
    izba_value = df_dict_izby.loc[df_dict_izby['IzbaId'] == i]['IzbaNazwa'].values[0]
    logCzas=datetime.now()
    logIzba=izba_value
    print(datetime.now()," Start-", izba_value)

    for j in driver.find_elements_by_class_name('wybor'):
        SadNazwa =j.text
        if SadNazwa[:3]=='Sąd' and SadNazwa.count("Sąd")==1:
            sady = {
                'IzbaId': i,
                'SadNazwa':SadNazwa
                    }
            dict_sady.append(sady)
    df_dict_sady = pd.DataFrame(dict_sady)

    for a in driver.find_elements_by_class_name('ListaKancelari'):
        for b in a.find_elements_by_tag_name("li"):
            sad_rej_temp = b.find_element_by_class_name('SadRejonowyC')
            sad_rej = sad_rej_temp.text
            for c, d in zip(b.find_elements_by_class_name('NazwaKancelarii'),b.find_elements_by_class_name('kancelaria_tresc')):
                komornik_nazwa = " ".join((c.text).split())
                dane_temp = " ".join((d.text).split())
                komornik_adres = dane_temp[0:dane_temp.index("Mapa")]

                if dane_temp.find(":") != -1:
                    komornik_konto = rachunek(dane_temp[dane_temp.index(": "):dane_temp.index("Izba")])

                if dane_temp.find("@") != -1 and dane_temp.find("mail ") != -1:
                    email_temp = dane_temp[dane_temp.index("mail "):][dane_temp[dane_temp.index("mail "):].index(" "):].lstrip()
                    komornik_email = email_temp[:email_temp.index(" ")]

                if dane_temp.find("@") == -1 and dane_temp.find("mail ") == -1:
                    komornik_email = "mail@techniczny.pl"
                if dane_temp.find("@") == -1 and dane_temp.find("mail ") != -1:
                    komornik_email = "mail@techniczny.pl"

                komornicy = {
                    'DataDanych':datetime.now(),
                    'KomornikRewir': sad_rej,
                    'KomornikNazwa': komornik_nazwa,
                    'KomornikAdres': komornik_adres,
                    'KomornikEmail': komornik_email,
                    'KomornikKonto': komornik_konto,
                    'KomornikIzba' : i
                            }
                dict_komornicy.append(komornicy)
        df_dict_komornicy = pd.DataFrame(dict_komornicy)                
    # wyświetlenie podsumowania
    liczbaKomornikow = df_dict_komornicy['KomornikIzba']==str(i)
    komornikowApelacja = df_dict_komornicy[liczbaKomornikow]['KomornikNazwa'].count()
    if komornikowApelacja==0:
        break
    else:
        print("Dodano "+str(komornikowApelacja)+" komorników.")
        print("-"*70)
        logDrill.append({'czas': str(logCzas), 'izba': logIzba, 'liczba': str(komornikowApelacja)})

    # czekam 5 sekund
    time.sleep(5)
    # zamykam przeglądarkę
    driver.quit()
# aktualizacja logu
df_logDrill = pd.DataFrame(logDrill)

# % eksport wyników do pliku, aktualizacja bazy nowymi rekordami %
print(datetime.now()," START - eksport wyników do plików csv")
df_dict_sady.to_csv(dataDanych+'_dict_sady.csv', index = False, header=True)
df_dict_komornicy.to_csv(dataDanych+'_dict_komornicy.csv', index = False, header=True)
df_dict_izby.to_csv(dataDanych+'dict_izby.csv', index = False, header=True)
df_logDrill.to_csv(dataDanych+'log.csv', index = False, header=True)
print(datetime.now()," STOP  - eksport wyników do plików csv")

dictKomornicy = pd.read_csv('../Data/dictKomornicy.csv')

wynik = df_dict_komornicy.merge(dictKomornicy, on='KomornikNazwa', how='outer')
wynik['DataDanych_y'].fillna(value=0, inplace=True)
filtr = wynik["DataDanych_y"] == 0
noweRekordy=wynik[filtr]['DataDanych_x'].to_frame()

if len(noweRekordy)>0:
    update = df_dict_komornicy.merge(noweRekordy, left_on='DataDanych', right_on='DataDanych_x', how='inner')
    update.drop('DataDanych_x', axis=1, inplace=True)
    print('Nowych rekordów: ',len(noweRekordy))
    dictKomornicy = dictKomornicy.append(update)
    dictKomornicy.to_csv('../Data/dictKomornicy.csv', index = False, header=True)
else:
    print('Brak nowych aktualizacji.')
    
# % wysłka maila %
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

emailfrom  = ""
emailto    = ""
fileToSend = '../Data/dictKomornicy.csv'
username   = "HubOn.pl@outlook.com"
password   = ""

msg = MIMEMultipart()
msg["From"] = emailfrom
msg["To"] = emailto
msg["Subject"] = "Aktualizacja bazy komorników"
msg.preamble = "W zalaczeniu aktualizacja bazy komornikow."

ctype, encoding = mimetypes.guess_type(fileToSend)
if ctype is None or encoding is not None:
    ctype = "application/octet-stream"

maintype, subtype = ctype.split("/", 1)

if maintype == "text":
    fp = open(fileToSend)
    # Note: we should handle calculating the charset
    attachment = MIMEText(fp.read(), _subtype=subtype)
    fp.close()
elif maintype == "image":
    fp = open(fileToSend, "rb")
    attachment = MIMEImage(fp.read(), _subtype=subtype)
    fp.close()
elif maintype == "audio":
    fp = open(fileToSend, "rb")
    attachment = MIMEAudio(fp.read(), _subtype=subtype)
    fp.close()
else:
    fp = open(fileToSend, "rb")
    attachment = MIMEBase(maintype, subtype)
    attachment.set_payload(fp.read())
    fp.close()
    encoders.encode_base64(attachment)
attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
msg.attach(attachment)

server = smtplib.SMTP('smtp-mail.outlook.com', 587)
server.starttls()
server.login(username,password)
server.sendmail(emailfrom, emailto, msg.as_string())
server.quit()
