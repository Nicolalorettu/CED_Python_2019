import xml.etree.ElementTree as ET
import subprocess
import pyautogui as p
import sys
import time
import os
import datetime as dt
import pandas as pd


def dldata():
    anno = 2018
    check = True
    icheck = True
    mesi = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre",
            "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre",
            "GENNAIO", "FEBBRAIO", "MARZO", "APRILE", "MAGGIO", "GIUGNO", "LUGLIO", "AGOSTO", "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"]
    while icheck:
        mese = input("Inserisci il mese (in lettere): ")
        while check:
            for i in mesi:
                if i == mese:
                    check = False
            if check == True:
                print("Mese inserito errato \n")
                mese = input("Inserisci il mese (in lettere): ")

        mese = mesi.index(mese) % 12 + 1

        check = True
        giorno = input("Inserisci il giorno: ")
        while check:
            try:
                date = dt.date(anno, int(mese), int(giorno))
                check = False
            except ValueError:
                print("Giorno inserito errato \n")
                giorno = input("Inserisci il giorno: ")

        i = dt.date(dt.datetime.now().year, dt.datetime.now().month, dt.datetime.now().day) - date
        if i.days < 20:
            return i.days, date.strftime("\%d_%m_%Y")
            icheck = False
        else:
            print("Non puoi scaricare una data più vecchia di 20 giorni")


def load(path):
    checkload = 0
    loading = True
    while loading:
        try:
            x, y = p.locateCenterOnScreen(path)
            loading = False
            return x, y
        except:
            checkload += 1
            print(checkload)
        time.sleep(3)
        if checkload == 30:
            loading = False
            quit()


def updmultiday(i):
    for j in range(0, 2):
        offsetx = 185 * j
        coord = [475 + offsetx, 515]
        # scarica un numero di giorni pari ad "i"
        oldestdate = dt.datetime.today() - dt.timedelta(i)
        month = oldestdate.month
        year = oldestdate.year
        if dt.date(year, month, 1).isoweekday() == 7:
            firstday = 0
        else:
            firstday = dt.date(year, month, 1).isoweekday()
        ndayfrow = 7 - firstday

        if dt.datetime.today().day <= i:
            p.click(475 + offsetx, 470)
            time.sleep(2)

        datefrom = dt.datetime.today() - dt.timedelta(i)
        datefrom = datefrom.day

        if datefrom <= ndayfrow:
            ycal = 0
        else:
            ycal = (datefrom - ndayfrow) // 7 + 1
        if dt.date(year, month, datefrom).isoweekday() == 7:
            xcal = 0
        else:
            xcal = dt.date(year, month, datefrom).isoweekday()
        xcoord = coord[0] + (25 * xcal)
        ycoord = coord[1] + (20 * ycal)
        time.sleep(2)
        p.click(xcoord, ycoord)


def tmc_conn():
    filenames = list()
    dfcols = list()
    dflist = list()
    dffinal = pd.DataFrame()
    dlgiorni, today = dldata()
    try:
        os.mkdir(r"c:" + today)
    except FileExistsError:
        None
    os.system("taskkill /f /t /im chrome.exe")
    keys = ["tab", "down", "space", "enter", "up", "right"]
    service = ["11", "12"]
    credlist = ["x0150576", "chrome.2018"]
    keypress = ["02", "52", "15", "31"]
    subprocess.Popen('"C:\Program Files\Google\Chrome\Application\chrome.exe" --new-window --ignore-certificate-errors --new-window --run-all-flash-in-allow-mode --test-type --args --start-maximized "https://paco.telecomitalia.local"', shell=True)
    time.sleep(20)
    try:
        x, y = p.locateCenterOnScreen(r'.\image\chrome ripristino pagine.png')
        p.click(x+180, y-10)
        time.sleep(1)
    except:
        None
    p.click(575, 400)
    p.press("enter")
    time.sleep(2)
    p.press("tab", interval=0.5)
    p.typewrite(credlist[0], interval=0.1)
    p.press("tab")
    p.typewrite(credlist[1], interval=0.1)
    p.press("enter")
    time.sleep(20)
    for j in range(0, len(keypress)):
        p.press(keys[int(keypress[j][:1])], presses=int(keypress[j][1:]), interval=0.2)
        time.sleep(0.5)
    x, y = load(r'.\image\paco2017_1280.png')
    x += 22
    p.click(x, y)
    time.sleep(1)
    updmultiday(dlgiorni)
    p.press(["tab", "space"], interval=0.5)
    x, y = load(r'.\image\paco Totale ENNOVA_1280.png')
    p.click(x, y)
    load(r'.\image\paco asa_1280.png')
    x, y = load(r'.\image\paco ennova_1280.png')
    p.click(x, y)
    load(r'.\image\paco bacino_1280.png')
    x, y = load(r'.\image\paco export_1280.png')
    p.click(x, y)
    time.sleep(10)
    filename = "\gruppo.xls"
    p.typewrite(r"c:" + today + filename)
    p.press("enter")
    time.sleep(10)
    tree = ET.parse(r"c:" + today + filename)
    root = tree.getroot()
    i = 0
    for ws in root:
        for table in ws:
            for row in table:
                for cell in row:
                    for data in cell:
                        if i == 25:
                            dfcols.append(data.text)
                        if i > 25:
                            dflist.append(data.text)
                i += 1
                if i == 26:
                    df = pd.DataFrame(columns=dfcols)
                if i > 26:
                    df = df.append(pd.Series(dflist, index=dfcols), ignore_index=True)
                dflist = []
    ngroups = df["Sedi"].count()-1
    total = df[-1:]
    for i in range (0, ngroups):
        p.press("tab", presses=2, interval=0.5)
        p.press("down", presses=i+1, interval=0.5)
        time.sleep(2)
        x, y = load(r'.\image\paco enter group_1280.png')
        p.click(x+100, y)
        x, y = load(r'.\image\paco load sede_1280.png')
        x, y = load(r'.\image\paco export_1280.png')
        p.click(x, y)
        time.sleep(10)
        filename = "\\" + df["Sedi"][i] + ".xls"
        filenames.append(filename)
        p.typewrite(r"c:" + today + filename)
        p.press("enter")
        time.sleep(10)
        x, y = load(r'.\image\paco indietro_1280.png')
        p.click(x, y)
        time.sleep(10)
    for filename in filenames:
        dfcols = list()
        tree = ET.parse(r"c:" + today + filename)
        root = tree.getroot()
        i = 0
        for ws in root:
            for table in ws:
                for row in table:
                    for cell in row:
                        for data in cell:
                            if i == 25:
                                dfcols.append(data.text)
                            if i > 25:
                                dflist.append(data.text)
                    i += 1
                    if i == 26:
                        df = pd.DataFrame(columns=dfcols)
                    if i > 26:
                        df = df.append(pd.Series(dflist, index=dfcols), ignore_index=True)
                    dflist = []
        df.drop(max(df.index), inplace=True)
        dffinal = pd.concat([dffinal, df], ignore_index=True)
        try:
            os.remove(r"c:" + today + filename)
        except FileNotFoundError:
            None
    total.columns = dfcols
    dffinal = pd.concat([dffinal, total], ignore_index=True)
    dffinal.to_excel(r"c:" + today + "\Paco TMC.xls", index=False)
    os.remove(r"c:" + today + "\gruppo.xls")


tmc_conn()
