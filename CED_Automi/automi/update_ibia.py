import pyperclip as pc
import pyautogui as p
import time
import subprocess
import datetime as dt
import var.links_paths as lp
import var.sql as sql
import functions.mailsender as m
import functions.check_connect as chc
import os
import sys
import ast
import pandas as pd

passo = 21
basex = 1106
basey = 294
passoxdrag = 200
increment = 6
increment2 = 2
safecount = 0
mailcount = 0
passomenu = 19

crs = sql.createCursor()
cnx = sql.cnx

sys.path.insert(0, r"../.")

import var.sql as sql


def start_connection():
    global vpncheck
    crs = sql.createCursor()
    query = sql.cred_query("ALL")
    crs.execute(query)
    credlist = crs.fetchone()
    vpncheck = chc.tim_vpn("Ibia")
    if vpncheck:
        loginIbia(credlist)


def loginIbia(credlist):
    global safecount
    os.system("taskkill /f /t /im chrome.exe")
    subprocess.Popen('"%s" --args --start-maximized --ignore-certificate-errors --auto-open-devtools-for-tabs --test-type --remote-debugging-port=9222 "%s"' % (lp.pathchrome, lp.Ibia))
    loginprocess = True
    while loginprocess:
        time.sleep(15)
        if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\loginibia.png'):
            p.typewrite(credlist[0], interval=0.5)
            time.sleep(0.2)
            p.press("tab")
            time.sleep(0.2)
            p.typewrite(credlist[1], interval=0.5)
            time.sleep(0.2)
            p.press("enter")
            loginprocess = False
        else:
            safecount += 1
            if safecount == 3:
                errors = open(lp.errorlogpath, "a")
                errors.write("[ %s ] Impossibile recuperare dati su Paco - Sito Irraggiungibile o connessione KO\n\r"
                             % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                errors.close()
                m.mail_error("Ibia SOS")
            else:
                time.sleep(5)
                loginIbia(credlist)
                loginprocess = True


def downloadxml_groups():
    response = "RenderWebPartContent"
    time.sleep(3)
    p.moveTo(basex, basey)
    time.sleep(0.2)
    p.scroll(2100)
    check = True
    x = 0
    while check:
        x += 1
        finaly = basey + (passo*x)
        if finaly <= 777:
            time.sleep(0.2)
            p.moveTo(basex, finaly)
        else:
            finaly = 777
            p.press("down", presses=2, interval=0.2)
            time.sleep(0.2)
            p.moveTo(basex, finaly)
        time.sleep(0.2)
        p.dragRel(passoxdrag, 0, 1, button='left')
        p.keyDown('ctrl')
        p.press('c')
        p.keyUp('ctrl')
        dragreal = pc.paste()
        if dragreal == response:
            p.rightClick()
            time.sleep(0.2)
            p.press("down", presses=4, interval=0.2)
            time.sleep(0.2)
            p.press("right", interval=0.2)
            time.sleep(0.2)
            p.press("down", presses=3, interval=0.2)
            time.sleep(0.2)
            p.press("enter")
            filefromcopy = str(pc.paste())
            if filefromcopy.find("ENNOVA_CA") > 0:
                check = False
    return filefromcopy


def downloadxmllast():
    global safecount
    global mailcount
    check = True
    while check:
        renderweb = "RenderWebPartContent"
        time.sleep(2)
        p.click(1104, 780)
        time.sleep(2)
        p.dragRel(passoxdrag, 0, 1, button='left')
        time.sleep(0.2)
        p.keyDown('ctrl')
        p.press('c')
        p.keyUp('ctrl')
        dragreal = pc.paste()
        if dragreal == renderweb:
            p.rightClick()
            time.sleep(0.2)
            p.press("down", presses=4, interval=0.2)
            time.sleep(0.2)
            p.press("right", interval=0.2)
            time.sleep(0.2)
            p.press("down", presses=3, interval=0.2)
            time.sleep(0.2)
            p.press("enter")
            cleanerIBIA(str(pc.paste()))
            check = False
        else:
            safecount += 1
            if safecount == 10:
                errors = open(lp.errorlogpath, "a")
                errors.write("[ %s ] Sito Web Bloccato, riavvio automa %s di 3 tentativi in corso\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), (mailcount)))
                errors.close()
                crs.close()
                mailcount += 1
                if mailcount == 3:
                    errors = open(lp.errorlogpath, "a")
                    errors.write("[ %s ] Impossibile recuperare dati su Paco - Sito Irraggiungibile o connessione KO\n\r"
                                 % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    errors.close()
                    m.mail_error("Ibia")
            check = True


def segnalazione_FE(groupdict):
    global increment
    global increment2
    grouplist = list(groupdict)
    time.sleep(3)
    test = True
    while test:
        if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint.png'):
            x, y = p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint.png')
            p.click(x, y)
            test = False
        else:
            if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint2.png'):
                x, y = p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint2.png')
                p.click(x, y)
                test = False
    time.sleep(0.2)
    p.press("tab", presses=29, interval=0.5)
    time.sleep(0.2)
    p.press("enter")
    time.sleep(0.2)
    p.doubleClick(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\selectall.png'), button='left', interval=0.5)
    time.sleep(0.2)
    for i in range(0, len(grouplist)):
        p.click(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\filtro.png'), button='left', interval=0.5)
        time.sleep(0.2)
        p.press("tab", presses=4, interval=0.5)
        time.sleep(0.2)
        if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\Selection.png'):
            p.press("enter")
        time.sleep(0.2)
        p.press("tab", presses=(i+1)*2, interval=0.5)
        p.press("enter")
        for j in range(0, groupdict[grouplist[i]]):
            p.press("tab", presses=j+2, interval=0.5)
            time.sleep(0.2)
            p.press("space")
            time.sleep(0.2)
            p.click(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\Ibiaok.png'))
            time.sleep(2)
            p.moveTo(basex, basey)
            time.sleep(0.2)
            p.scroll(-2000)
            time.sleep(0.5)
            downloadxmllast()
            time.sleep(0.2)
            test = True
            while test:
                if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint.png'):
                    x, y = p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint.png')
                    p.click(x, y)
                    test = False
                else:
                    if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint2.png'):
                        x, y = p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint2.png')
                        p.click(x, y)
                        test = False
            time.sleep(0.2)
            p.press("tab", presses=29, interval=0.5)
            time.sleep(0.2)
            p.press("enter")
            time.sleep(0.2)
            p.doubleClick(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\selectall.png'), button='left', interval=0.5)
            time.sleep(0.2)
            p.click(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\filtro.png'), button='left', interval=0.5)
            time.sleep(0.2)
            p.press("tab", presses=4, interval=0.5)
            time.sleep(0.2)
            p.press("tab", presses=(i+1)*2, interval=0.5)
            time.sleep(0.2)
            j = j+2
        j = 0
        time.sleep(0.2)
        p.press("enter")
        i = i+1



        # p.press("tab", presses=groupdict[grouplist[i]], interval=0.5)

def date_time():
    time.sleep(2)
    p.press("f5")
    time.sleep(10)
    p.click(405, 323)
    time.sleep(0.2)
    p.doubleClick(404, 545, button='left', interval=0.5)
    time.sleep(0.5)
    p.click(405, 352)
    time.sleep(0.2)
    crs = sql.createCursor()
    time.sleep(5)
    a = dt.datetime.now().strftime("%d %m %Y")[3:6]
    crs.execute("SELECT name FROM ced_month_tr WHERE number = %s" % a)
    mese = crs.fetchone()
    giorno = dt.datetime.now().strftime("%d %m %Y")[:3]
    anno = dt.datetime.now().strftime("%d %m %Y")[5:]
    p.typewrite(giorno + mese[0] + anno, interval=0.1)
    time.sleep(0.2)
    p.press("enter")
    time.sleep(2)
    p.press("tab", presses=2, interval=0.5)
    time.sleep(0.5)
    p.press("space")
    time.sleep(0.5)
    p.doubleclick(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\Ibiaok.png'), button='left', interval=0.5)


def navigate_on_ibia():
    keys = ["tab", "space", "enter"]
    groups = {"Bfonia": ["05", "21", "04", "21", "02", "11"],
              "BADSL": ["05", "21", "02", "21", "02", "11"],
              "Bfibra": ["05", "21", "02", "21", "04", "11"],
              "Hfonia": ["07", "21", "04", "21", "02", "11", "03", "21", "02", "11"],
              "HADSL": ["07", "21", "06", "21", "02", "11", "07", "21", "02", "11"],
              "Hfibra": ["07", "21", "06", "21", "04", "11"]}
    groupsback = {"Bfonia": ["09", "21", "04", "21"],
                  "BADSL": ["07", "21", "02", "21"],
                  "Bfibra": ["07", "21", "02", "21"],
                  "Hfonia": ["016", "21", "05", "21", "04", "21"],
                  "HADSL": ["022", "21", "09", "21", "06", "21"],
                  "Hfibra": ["013", "21", "06", "21"]}
    listajobs = list(groups)
    listavalues = list(groups.values())
    listavaluesback = list(groupsback.values())
    i = 0
    for i in range(0, len(listajobs)):
        time.sleep(3)
        test = True
        while test:
            if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint.png'):
                x, y = p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint.png')
                p.click(x, y)
                test = False
            else:
                if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint2.png'):
                    x, y = p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint2.png')
                    p.click(x, y)
                    test = False
        time.sleep(0.2)
        p.press("tab", presses=28, interval=0.5)
        time.sleep(0.2)
        p.press("enter")
        time.sleep(0.2)
        p.doubleClick(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\selectall.png'), button='left', interval=0.5)
        time.sleep(0.2)
        p.click(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\filtro.png'), button='left', interval=0.5)
        time.sleep(0.2)
        j = 0
        for j in range(0, len(listavalues[i])):
            p.press((keys[int(listavalues[i][j][:1])]), presses=int(listavalues[i][j][1:]), interval=0.5)
            time.sleep(0.2)
        p.click(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\Ibiaok.png'))
        # adatta la funzione segnalazione FE e mettila qua!
        test = True
        while test:
            if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint.png'):
                x, y = p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint.png')
                p.click(x, y)
                test = False
            else:
                if p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint2.png'):
                    x, y = p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\sharepoint2.png')
                    p.click(x, y)
                    test = False
        time.sleep(0.2)
        p.press("tab", presses=28, interval=0.5)
        time.sleep(0.2)
        p.press("enter")
        time.sleep(0.2)
        p.doubleClick(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\selectall.png'), button='left', interval=0.5)
        time.sleep(0.2)
        p.click(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\filtro.png'), button='left', interval=0.5)
        time.sleep(0.2)
        m = 0
        for m in range(0, len(listavaluesback[i])):
            if m == 2:
                p.keyDown("shift")
            p.press((keys[int(listavaluesback[i][m][:1])]), presses=int(listavaluesback[i][m][1:]), interval=0.5)
            time.sleep(0.2)
        p.keyUp("shift")
        p.click(p.locateCenterOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\Ibiaok.png'))


def dictionary_grouplist(groups):
    dfgruppi = pd.DataFrame()
    grouplen = dict()
    data = groups
    data = data[data.find('"Members"')+10:data.find(',"More"')]
    data = data.replace("true", "True")
    data = data.replace("false", "False")
    dictgruppi = ast.literal_eval(data)
    for i in range(0, len(dictgruppi)):
        dfgruppi = pd.concat([dfgruppi, pd.DataFrame(dictgruppi[i], index=[0])], axis=0, sort=False, ignore_index=True)
    for i in range(0, len(dfgruppi.index)):
        if dfgruppi["Lv"][i] < 3:
            check = dfgruppi["Tx"][i]
        dfgruppi.replace(dfgruppi["Id"][i], check, inplace=True)
    grouplist = dfgruppi[dfgruppi["Lv"] == 2]["Id"].tolist()
    for j in range(0, len(grouplist)):
        i = 0
        for k in range(0, len(dfgruppi[dfgruppi["Lv"] == 3]["Id"].tolist())):
            if grouplist[j] == dfgruppi[dfgruppi["Lv"] == 3]["Id"].tolist()[k]:
                i += 1
        grouplen[grouplist[j]] = i
    return grouplen

#                                                 #
# Tutte le funzioni fanno return ad un valore che #
# viene utilizzato nella funzione più esterne     #
#                                                 #

def update_ibia():
    start_connection()
    groups = downloadxml_groups()
    grouplen = ASCupdIBIA(groups)
    segnalazione_FE(grouplen)
