import xml.etree.ElementTree as ET
import functions.check_connect as chc
import functions.mailsender as m
import subprocess
import pyautogui as p
import sys
import time
import os
import datetime as dt
import functions.fromfiletosql_ivr as ffq
import functions.check_connect as cc
import pandas as pd

sys.path.insert(0, r"../.")

import var.links_paths as lp
import var.sql as sql

crs = sql.createCursor()
cnx = sql.cnx


def load(path, checkload=0):
    loading = True
    while loading:
        try:
            x, y = p.locateCenterOnScreen(path)
            loading = False
            return x, y
        except:
            checkload += 1
        time.sleep(1)
        if checkload == 60:
            loading = False
            quit()


def updmultiday(i = 1):
    for j in range(0, 2):
        offsetx = 225 * j
        coord = [760 + offsetx, 540]
        # scarica un numero di giorni pari ad "i"
        oldestdate = dt.datetime.today() - dt.timedelta(i)
        month = oldestdate.month
        year = oldestdate.year
        ndayfrow = 8 - dt.date(year, month, 1).isoweekday()

        if dt.datetime.today().day <= i:
            p.click(750 + offsetx, 480)
            time.sleep(2)

        datefrom = dt.datetime.today() - dt.timedelta(i)
        datefrom = datefrom.day

        ycal = (datefrom - ndayfrow) // 7 + 1
        xcal = dt.date(year, month, datefrom).isoweekday()
        xcoord = coord[0] + (28 * xcal)
        ycoord = coord[1] + (22 * ycal)
        time.sleep(2)
        p.click(xcoord, ycoord)


def tmc_conn():
    filenames = list()
    dfcols = list()
    dflist = list()
    dffinal = pd.DataFrame()
    today = (dt.datetime.now() - dt.timedelta(1)).strftime("\%d_%m_%Y")
    try:
        os.mkdir(r"c:" + today)
    except FileExistsError:
        None
    os.system("taskkill /f /t /im chrome.exe")
    keys = ["tab", "down", "space", "enter", "up", "right"]
    service = ["11", "12"]
    pacolink = (lp.PACOlink)
    query = sql.cred_query("ALL")
    crs.execute(query)
    credlist = crs.fetchone()
    keypress = ["02", "52", "15", "31"]
    subprocess.Popen('"%s" --new-window --ignore-certificate-errors --new-window --run-all-flash-in-allow-mode --test-type --args --start-maximized "%s"' % (lp.pathchrome, pacolink))
    x, y = load(r'.\image\chrome ripristino pagine.png', checkload=55)
    p.click(x+180, y-10)
    time.sleep(10)
    p.click(575, 400)
    time.sleep(1)
    p.press("tab")
    time.sleep(1)
    p.press("enter")
    time.sleep(15)
    p.click(575, 400)
    p.press("enter")
    time.sleep(2)
    p.press("tab", presses=2, interval=0.5)
    p.typewrite(credlist[0], interval=0.1)
    p.press("tab")
    p.typewrite(credlist[1], interval=0.1)
    p.press("enter")
    time.sleep(20)
    for j in range(0, len(keypress)):
        p.press(keys[int(keypress[j][:1])], presses=int(keypress[j][1:]), interval=0.2)
        time.sleep(0.5)
    x, y = load(r'.\image\paco2017.png')
    x += 22
    p.click(x, y)
    time.sleep(1)
    updmultiday()
    p.press(["tab", "space"], interval=0.5)
    x, y = load(r'.\image\paco Totale ENNOVA.png')
    p.click(x, y)
    load(r'.\image\paco asa.png')
    x, y = load(r'.\image\paco ennova.png')
    p.click(x, y)
    load(r'.\image\paco bacino.png')
    x, y = load(r'.\image\paco export.png')
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
        x, y = load(r'.\image\paco enter group.png')
        p.click(x+100, y)
        x, y = load(r'.\image\paco load sede.png')
        x, y = load(r'.\image\paco export.png')
        p.click(x, y)
        time.sleep(10)
        filename = "\\" + df["Sedi"][i] + ".xls"
        filenames.append(filename)
        p.typewrite(r"c:" + today + filename)
        p.press("enter")
        time.sleep(10)
        x, y = load(r'.\image\paco indietro.png')
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
