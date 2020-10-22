#!/usr/bin/python
# -*- coding: utf-8 -*-

import requests
import re
from bs4 import BeautifulSoup
import time
import datetime

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pyodbc
import requests
import traceback

from PIL import Image
from io import BytesIO
import webbrowser

import pandas as pd

import random

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
    start0 = time.time()

    user = input("Vnesi uporabnika:") # valerija1999
    password = input("Vnesi password:") # 19Valerija73!

    # dict_sif_tip = {"04": "Letno poročilo gospodarske družbe in zadruge",
    #                 "05": "Letno poročilo samostojnega podjetnika",
    #                 "06": "Revidirano letno poročilo gospodarske družbe in zadruge",
    #                 "07": "Konsolidirano letno poročilo gospodarske družbe in zadruge",
    #                 "10": "Letno poročilo s preiskanimi računovodskimi izkazi",
    #                 "20": "Letno poročilo društva",
    #                 "21": "Revidirano letno poročilo društva"}

    dict_sif_tip = {"01": "Letno poročilo",
                    "02": "Revidirano letno poročilo",
                    "03": "Konsolidirano letno poročilo",
                    "04": "Letno poročilo s preiskanimi računovodskimi izkazi"}

    print ("Razpolozljivi tipi porocil so:")
    print (dict_sif_tip)
    print("\n")
    test1 = False
    test2 = False
    test3 = False

    while test1 == False:
        sif_tip = input("Vnesi sifro tipa porocila (v obliki npr. 04): ")
        print("\n")
        print("Vnesli ste naslednji tip: " + sif_tip + " - " + dict_sif_tip[sif_tip])
        preverba = input("Ali je vas vnos pravilen (da/ne): ")
        if preverba == "da":
            test1 = True
            print("\n")
            print("Izvedlo se bo snemanje za: " + sif_tip + " - " + dict_sif_tip[sif_tip])
        if sif_tip == "01":
            test2 = True

    while test3 == False:
        leto = input("Vnesi leto (v obliki npr. 2018): ")
        print("\n")
        print("Vnesli ste naslednje leto: " + leto)
        preverba = input("Ali je vas vnos pravilen (da/ne): ")
        if preverba == "da":
            test3 = True
            print("\n")
            print("Izvedlo se bo snemanje za: " + leto)



    df_mat_st = pd.read_excel("C:\\0-Bilance\\0-Snemanje bilanc\\Snemanje LP-ajpes\\MS_za_snemanje_PDF.xlsx")
    #df_mat_st = pd.read_excel("c:\\drustva\\Mat-st-vnos.xlsx", header=None)
    df_mat_st["Datum_objave"] = df_mat_st["Datum_objave"].dt.strftime("%d.%m.%Y")

    #odstranim dvojnike
    df_mat_st.drop_duplicates(keep="first",inplace=True)

    lista_mat_st = df_mat_st.values.tolist()


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
    soup_prijava = BeautifulSoup(r_sub_prijava.text)
    soup_prijava_find = soup_prijava.find(text=re.compile('"error": 0'))

    if not soup_prijava_find:
        print("Napaka pri prijavi!")
    else:
        print("Prijava uspešna!")

    stevec = 0

    lista_glavna = []

    for i in lista_mat_st:

        try:

            #mat_st_no = 5539978000
            mat_st_no = i[0]
            #datum_objave = "09.09.2019"
            datum_objave = i[1]

            mat_st_str = str(mat_st_no)

            url_jolp = "https://www.ajpes.si/jolp/podjetje.asp?maticna=" + mat_st_str

            r_sub_jolp = s.get(url_jolp, verify=False)

            r_sub_jolp.encoding = "utf-8"
            soup_jolp = BeautifulSoup(r_sub_jolp.text)

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
                    img.save("C:\\0-Bilance\\0-Snemanje bilanc\\Snemanje LP-ajpes\\captcha.jpg", format="JPEG")
                    webbrowser.open('C:\\0-Bilance\\0-Snemanje bilanc\\Snemanje LP-ajpes\\captcha.jpg')

                    captcha_url = "https://www.ajpes.si/jolp/podjetje.asp?maticna=" + mat_st_str
                    koda = input("Vpiši Captcho:")
                    kljukica = "1"

                    payload_cap = {'jolp_koda': koda, 'izjava_strinjam': kljukica}

                    s.post(captcha_url, data=payload_cap, verify=False)

                    r_sub_jolp = s.get(url_jolp, verify=False)
                    r_sub_jolp.encoding = "utf-8"
                    soup_jolp = BeautifulSoup(r_sub_jolp.text)

                    captcha = soup_jolp.find_all(text=re.compile("Za dostop do podatkov"))

                    if captcha:
                        print("Vpisali ste napačno Captcho! Prosim, ponovite vpis!")
                    else:
                        print("Vpisali se pravilno Captcho!")
                        print("\n")
                        test = False

            naziv = soup_jolp.find("h4").text

            os_string = soup_jolp.find(text=re.compile("Vrsta poročila"))
            tabela = os_string.find_parent("table")

            os_string_2 = tabela.find_all("td", text=re.compile(datum_objave))

            zap_st = 1

            for n in os_string_2:
                lista_delna = []
                tabela_vrstica = n.find_parent("tr")
                lista_vrstica = tabela_vrstica.find_all("td")

                if lista_vrstica[0].find(text=re.compile(dict_sif_tip[sif_tip])):

                    for m in lista_vrstica:
                        vrstica = m.text.strip().strip('\t\n\r')
                        lista_delna.append(vrstica)
                    link = tabela_vrstica.find("a").get("href")
                    #lista_delna.append(link)
                    if zap_st == 1:
                        if test2 == True:
                            file_name = mat_st_str[:7]
                        else:
                            file_name = mat_st_str[:7] + "-" + leto
                    else:
                        if test2 == True:
                            file_name = mat_st_str[:7] + "-" + str(zap_st)
                        else:
                            file_name = mat_st_str[:7] + "-" + leto + "-" + str(zap_st)

                    url_file = "https://www.ajpes.si/jolp/" + link
                    print(url_file)

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
                    print(file_total)

                    # f = open("C:\\AjpesDokumenti\\insolv\\" + file_total, "wb")
                    f = open("C:\\0-Bilance\\0-Snemanje bilanc\\Snemanje LP-ajpes\\" + file_total, "wb")
                    f.write(r_file.content)

                    f.close()

                    lista_delna.append(file_total)
                    lista_delna.insert(0, mat_st_str)
                    lista_delna.insert(0, naziv)

                    lista_glavna.append(lista_delna)
                    print("\n")
                    zap_st = zap_st + 1

        except Exception as e:
            print(e)
            a = traceback.format_exc()
            print(a)

            tip = "Napaka - AJPES - JOLP, snemanje PDF-ov - posamicne napake"
            text = "Napaka - AJPES - JOLP, snemanje PDF-ov - mat. st.: " + mat_st_str + "\r\r\n" + "Napaka: " + str(e) + "\n" + "\nVrstica: " + str(a)
            e_mail("dba@ebonitete.si", "pravnapisarna@prvafina.si", tip, text, "mail.prvafina.si")
            e_mail("dba@ebonitete.si", "zoran.pesic@prvafina.si", tip, text, "mail.prvafina.si")

            pass

    col_names = ["Naziv", "Maticna_dolga", "Vrsta_porocila", "Za_leto/obdobje", "Datum_javne_objave", "Verzija", "Dokument", "Ime_datoteke"]
    df = pd.DataFrame(lista_glavna, columns=col_names)
    df.index += 1
    file_name = "C:\\0-Bilance\\0-Snemanje bilanc\\Snemanje LP-ajpes\\obdelane_MS.csv"
    df.to_csv(file_name, sep=';', encoding='windows-1250')


    print("Downloadanje zakljuceno!")



except Exception as e:
    print(e)
    a = traceback.format_exc()
    print(a)

    tip = "Napaka - AJPES - JOLP, snemanje PDF-ov - generalna napaka"
    text = "Napaka: " + str(e) + "\n" + "\nVrstica: " + str(a)
    e_mail("dba@ebonitete.si", "pravnapisarna@prvafina.si", tip, text, "mail.prvafina.si")
    e_mail("dba@ebonitete.si", "zoran.pesic@prvafina.si", tip, text, "mail.prvafina.si")