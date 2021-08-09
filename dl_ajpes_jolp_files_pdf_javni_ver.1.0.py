#!/usr/bin/python
# -*- coding: utf-8 -*-

import re
from bs4 import BeautifulSoup
import time

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests
import traceback

from PIL import Image
from io import BytesIO
import webbrowser

import pandas as pd
import datetime


# izklopim opozorilo, da ni varne SSL povezave
requests.packages.urllib3.disable_warnings()

def e_mail(me, you, tip, text, smtp_server):
    # me == my email address
    # you == recipient's email address


    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = tip
    msg['From'] = me
    msg['To'] = you

    # Create the body of the message (a plain-text ).

    # Record the MIME types of both parts - text/plain a.
    part1 = MIMEText(text, 'plain')


    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, is best and preferred.
    msg.attach(part1)


    # Send the message via local SMTP server.
    s = smtplib.SMTP(smtp_server)
    # sendmail function takes 3 arguments: sender's address, recipient's address
    # and message to send - here it is sent as one string.
    s.sendmail(me, you, msg.as_string())
    s.quit()


#####


try:
    # naredim štoparico
    start = time.time()

    #user = "valerija1999"
    #password = "19Valerija73!"

    user = input("Vnesi uporabnika:") # valerija1999
    password = input("Vnesi password:") # 19Valerija73!

    test = False

    while test == False:
        leto_od = input("Vnesi leto od (v obliki npr. 2018): ")
        leto_do = input("Vnesi leto do (v obliki npr. 2018): ")
        print("\n")
        print(f"Vnesli ste naslednji razpon let: od {leto_od} do {leto_do}" )
        preverba = input("Ali je vas vnos pravilen (da/ne): ")
        if preverba == "da":
            test = True
            print("\n")
            print(f"Izvedlo se bo snemanje naslednji razpon let: od {leto_od} do {leto_do}" )
            razpon = list(range(int(leto_od), int(leto_do) + 1))

    print("Snemam dokumente iz letnih poročil - Poslovna porocila s pojasnili")


    df_mat_st = pd.read_excel("C:\\0-Bilance\\0-Snemanje bilanc\\Snemanje PDF-Javni\\MS_za_snemanje_PDF.xlsx")
    #df_mat_st = pd.read_excel("MS_za_snemanje_PDF.xlsx")

    #odstranim dvojnike
    df_mat_st.drop_duplicates(keep="first",inplace=True)

    lista_mat_st = df_mat_st["MS_dolga"].tolist()


    s = requests.session()
    header = {"User_Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0",
              "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
              "Accept-Language": "sl,en-GB;q=0.7,en;q=0.3", "Accept-Encoding": "gzip, deflate",
              "Connection": "keep-alive"}

    s.headers = header

    payload = {'uporabnik': user, 'geslo': password}
    url_log = "https://www.ajpes.si/MDScripts/ajax.asp?method=checkuser"
    r_sub_prijava = s.post(url_log, data=payload, verify=False)
    r_sub_prijava.encoding = "utf-8"
    soup_prijava = BeautifulSoup(r_sub_prijava.text,features="lxml")
    soup_prijava_find = soup_prijava.find(text=re.compile('"error": 0'))

    if not soup_prijava_find:
        print("Napaka pri prijavi!")
    else:
        print("Prijava uspešna!")

    obdelane = []

    for count, ms in enumerate(lista_mat_st):

        try:

            #lista_mat_st = 2482851000
            mat_st_no = ms

            mat_st_str = str(ms)

            url_jolp = "https://www.ajpes.si/jolp/podjetje.asp?maticna=" + mat_st_str

            r_sub_jolp = s.get(url_jolp, verify=False)

            r_sub_jolp.encoding = "utf-8"
            soup_jolp = BeautifulSoup(r_sub_jolp.text, features="lxml")

            captcha = soup_jolp.find_all(text=re.compile("Za dostop do podatkov"))

            if captcha:

                test = True

                while test == True:

                    print("Potrebno je vpisati Captcho!")

                    link = soup_jolp.find("img", class_="captcha")
                    src_add = link["src"]
                    src_full = "https://www.ajpes.si/" + src_add[3:]

                    image = s.get(src_full)
                    img = Image.open(BytesIO(image.content))
                    img.save("captcha.jpg", format="JPEG")
                    webbrowser.open('captcha.jpg')

                    captcha_url = "https://www.ajpes.si/jolp/podjetje.asp?maticna=" + mat_st_str
                    koda = input("Vpiši Captcho:")
                    kljukica = "1"

                    payload_cap = {'jolp_koda': koda, 'izjava_strinjam': kljukica}

                    s.post(captcha_url, data=payload_cap, verify=False)

                    r_sub_jolp = s.get(url_jolp, verify=False)
                    r_sub_jolp.encoding = "utf-8"
                    soup_jolp = BeautifulSoup(r_sub_jolp.text, features="lxml")

                    captcha = soup_jolp.find_all(text=re.compile("Za dostop do podatkov"))

                    if captcha:
                        print("Vpisali ste napačno Captcho! Prosim, ponovite vpis!")
                    else:
                        print("Vpisali se pravilno Captcho!")
                        print("\n")
                        test = False

            print(f"Snemam zap. st. {count + 1} - MS: {ms}")

            # Pridobim naziv subjekta
            naziv = soup_jolp.find("h4").text

            # Poiscem tabelo in sparsam potrebne podatke
            tabela = soup_jolp.find("table", class_="table table-bordered")
            rows = tabela.find_all("tr")[1:]

            porocila = {}
            for i in rows:
                row = i.find_all("td")
                leto = int(row[1].text)
                porocilo = row[4].find("a", string=re.compile("Poslovno poročilo"))
                if porocilo:
                    link = "https://www.ajpes.si/jolp/" + porocilo["href"]
                    ime_datoteke = porocilo.text
                    porocila[leto] = [link, ime_datoteke]

            # Naredim novi dict z letnicami (keys) iz razpona
            porocila_v_razponu = { leto: porocila[leto] for leto in razpon if leto in porocila}

            for key, value in porocila_v_razponu.items():
                file_name = f"{mat_st_str[:7]}-{key}"
                print(f"   leto {key} - dokument: {value[1]}")

                url_file = value[0]

                r_file = s.get(url_file, stream=True, verify=False)

                hed = r_file.headers
                val_hed = hed.values()

                if "application/pdf" in val_hed:
                    file_ext = ".pdf"
                elif "application/PDF" in val_hed:
                    file_ext = ".pdf"
                elif "application/tif" in val_hed:
                    file_ext = ".tif"
                elif "application/TIF" in val_hed:
                    file_ext = ".tif"
                elif "application/tiff" in val_hed:
                    file_ext = ".tiff"
                elif "application/TIFF" in val_hed:
                    file_ext = ".tiff"
                else:
                    print("ERROR - extensions")

                file_total = file_name + file_ext
                #print(file_total)

                #f = open(file_total, "wb")
                f = open("C:\\0-Bilance\\0-Snemanje bilanc\\Snemanje PDF-Javni\\Downloads-temp\\" + file_total, "wb")
                f.write(r_file.content)

                f.close()

                obdelane.append([naziv, mat_st_str, key, file_total])




        except Exception as e:
            print(e)
            a = traceback.format_exc()
            print(a)

            tip = "Napaka - AJPES - JOLP, snemanje PDF-ov (Poslovna porocila s pojasnili) - posamicne napake"
            text = "Napaka - AJPES - JOLP, snemanje PDF-ov - mat. st.: " + mat_st_str + "\r\r\n" + "Napaka: " + str(e) + "\n" + "\nVrstica: " + str(a)
            e_mail("dba@ebonitete.si", "pravnapisarna@prvafina.si", tip, text, "mail.prvafina.si")
            e_mail("dba@ebonitete.si", "zoran.pesic@prvafina.si", tip, text, "mail.prvafina.si")

            pass


    col_names = ["Naziv", "Maticna_dolga", "Leto", "Ime_datoteke"]
    df = pd.DataFrame(obdelane, columns=col_names)
    df.index += 1
    #file_name = "obdelane_MS.csv"
    file_name = "C:\\0-Bilance\\0-Snemanje bilanc\\Snemanje PDF-Javni\\obdelane_MS.csv"
    df.to_csv(file_name, sep=';', encoding='windows-1250')

    print("Downloadanje zakljuceno!")

    # Ustavim stoparico
    stop = time.time()
    # Pridobim rezultat
    razlika = stop - start
    # Pretvorim sekunde od stopraice v bolj citljiv zapis
    pretvorba = str(datetime.timedelta(seconds=razlika))
    print(f'Potreben cas izvedbe: {pretvorba}\n')



except Exception as e:
    print(e)
    a = traceback.format_exc()
    print(a)

    tip = "Napaka - AJPES - JOLP, snemanje PDF-ov (Poslovna porocila s pojasnili) - generalna napaka"
    text = "Napaka: " + str(e) + "\n" + "\nVrstica: " + str(a)
    e_mail("dba@ebonitete.si", "pravnapisarna@prvafina.si", tip, text, "mail.prvafina.si")
    e_mail("dba@ebonitete.si", "zoran.pesic@prvafina.si", tip, text, "mail.prvafina.si")