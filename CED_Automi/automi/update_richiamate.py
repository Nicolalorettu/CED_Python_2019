import functions.check_connect as chc
import functions.mailsender as m
import subprocess
import pyautogui as p
import time
import os
import sys
import datetime as dt
import pychrome
import pyperclip as pc
import functions.fromfiletosql_richiamate as ffq
import var.links_paths as lp
import var.sql as sql

crs = sql.createCursor()
cnx = sql.cnx


wday = dict(zip("6012345", range(7)))
coord = (544, 530)
statuscode = 0
dizio = dict()
i = 0
t = 1
r = 0
basex = 1106
basey = 294
passo = 21


def request_will_be_sent(**kwargs):               # funzione lettura requests
    global r
    global risposta
    global dizio
    if r == 0:
        dizio = dict()
    risposta = (kwargs.get('request').get('url'))
    dizio[str(r)] = risposta
    print(dizio)
    r += 1
    return(risposta, r)


def loginpaco(credlist):
    global tab
    global t
    global r
    r = 0
    os.system("taskkill /f /t /im chrome.exe")
    time.sleep(5)
    subprocess.Popen('"%s" --args --start-maximized --ignore-certificate-errors --auto-open-devtools-for-tabs --test-type --remote-debugging-port=9222 "%s"' % (lp.pathchrome, lp.PACOlink))
    time.sleep(2)
    browser = pychrome.Browser()
    tab = browser.new_tab(url="https://10.104.100.140")
    tab.start()
    listtab = []
    listtab = browser.list_tab()
    browser.close_tab(listtab[1])
    tab.Network.enable()
    loginprocess = True
    oklocation = True
    matrlocation = True
    passwordlocation = True
    stuck = 0
    time.sleep(2)
    p.click(545, 750)
    time.sleep(2)
    p.click(325, 169)
    time.sleep(5)
    while loginprocess:
        oklocation = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\okbutton.png")
        if oklocation is None:
            oklocation = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\okbutton.png")
            oklocation = True
            stuck += 1
            time.sleep(5)
            if stuck == 10:
                errors = open(lp.errorlogpath, "a")
                errors.write("[ %s ] Sito Web Bloccato, riavvio automa %s di 3 tentativi in corso\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), (t)))
                errors.close()
                t += 1
                r = 0
                update_richiamate()
                if t == 3:
                    errors = open(lp.errorlogpath, "a")
                    errors.write("[ %s ] Impossibile recuperare dati su Paco - Sito Irraggiungibile o connessione KO\n\r"
                                 % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    errors.close()
                    m.mail_error("Paco - Richiamate")
        else:
            okx, oky = p.center(oklocation)
            time.sleep(2)
            p.click(okx, oky)
            time.sleep(1)
            oklocation = False
            matrlocation = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\matricola.png")
            if matrlocation is None:
                matrlocation = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\matricola.png")
                matrlocation = True
            else:
                matrx, matry = p.center(matrlocation)
                matrx = matrx + 40
                time.sleep(1)
                p.click(matrx, matry)
                time.sleep(1)
                p.typewrite(credlist[0], interval=0.2)
                time.sleep(2)
                matrlocation = False
                passwordlocation = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\password.png")
                if passwordlocation is None:
                    passwordlocation = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\password.png")
                    passwordlocation = True
                else:
                    pswx, pswy = p.center(passwordlocation)
                    pswx = pswx + 40
                    time.sleep(1)
                    p.click(pswx, pswy)
                    time.sleep(1)
                    p.typewrite(credlist[1], interval=0.2)
                    time.sleep(2)
                    passwordlocation = False
                    p.press("tab")
                    time.sleep(2)
                    p.press("space")
                    time.sleep(2)
                    loginprocess = False
                    errors = open(lp.accesslogpath, "a")
                    errors.write("[ %s ] Login Paco Effettuato con Successo\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    errors.close()
    return(tab)


def paco_ripetute(g, table):
    global t
    stuck = 0
    passox = 16
    passoxdrag = 400
    time.sleep(5)
    p.click(346, 197)                            # emptyclick
    openaccessshow = True
    pacoloading = True
    time.sleep(5)
    while openaccessshow:
        buttonshow = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\openaccess.png")
        if buttonshow is None:
            openaccessshow = True
            stuck += 1
            if stuck == 10:
                errors = open(lp.errorlogpath, "a")
                errors.write("[ %s ] Sito Web Bloccato, riavvio automa %s  di 3 tentativi in corso\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), (t)))
                errors.close()
                t += 1
                update_richiamate()
                if t == 3:
                    errors = open(lp.errorlogpath, "a")
                    errors.write("[ %s ] Impossibile recuperare dati su Paco - File non scaricato o illeggible\n\r"
                                 % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    errors.close()
                    m.mail_error("Paco - Richiamate")
                else:
                    pass
            else:
                pass
        else:
            btnx, btny = p.center(buttonshow)
            p.click(btnx, btny)
            time.sleep(2)
            p.press("down")
            time.sleep(2)
            p.press("enter")
            openaccessshow = False
            while pacoloading:
                time.sleep(2)
                if risposta == "https://10.104.100.140/Stats/OpenAccess/NuoveAbbRip/openAccess.aspx":
                    time.sleep(2)
                    h = 0
                    trova = True
                    while trova:
                        openaccess = p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\paper.png')
                        if openaccess is not None:
                            x, y = p.center(openaccess)
                            x2 = x + 8
                            time.sleep(5)
                            p.click(x2, y)
                            trova = False
                    ultimares = True
                    while ultimares:
                        y2 = y + (passo*h)
                        time.sleep(2)
                        p.click(x2, y2)
                        time.sleep(2)
                        check = p.pixel(x2, y2)
                        findblue = (33, 150, 243)
                        if check == findblue:
                            h += 1
                            ultimares = True
                        else:
                            y2 = y2 - passo
                            p.click(x2, y2)
                            ultimares = False
                    time.sleep(2)
                    headerclick = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\Headerslocation.png")
                    if headerclick is None:
                        headerclick = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\Headersclicked.png")
                    else:
                        headerx, headery = p.center(headerclick)
                        time.sleep(2)
                        p.click(headerx, headery)
                        time.sleep(1)
                    responselocation = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\responselocation.png")
                    time.sleep(2)
                    if responselocation is None:
                        errors = open(lp.accesslogpath, "a")
                        errors.write("[ %s ] Nessuna Response OnScreen\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                        errors.close()
                        responselocation = True
                    else:
                        buttresx, buttresy = p.center(responselocation)
                        time.sleep(2)
                        p.click(buttresx, buttresy)
                        time.sleep(2)
                        dragcheck = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\Iniziodrag.png")
                        if dragcheck is None:
                            errors = open(lp.accesslogpath, "a")
                            errors.write("[ %s ] Xml in Caricamento\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                            errors.close()
                            responselocation = True
                        else:
                            dragx, dragy = p.center(dragcheck)
                            dragx = dragx + passox
                            time.sleep(2)
                            p.click(dragx, dragy)
                            time.sleep(2)
                            dragreal = p.dragRel(passoxdrag, 0, 2, button='left')
                            p.keyDown('ctrl')
                            p.press('c')
                            p.keyUp('ctrl')
                            xmlcontrol = pc.paste()
                            if xmlcontrol is None or xmlcontrol != lp.xmlhead:
                                    errors = open(lp.accesslogpath, "a")
                                    errors.write("[ %s ] Nessun Xml Valido\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                                    errors.close()
                                    p.click(dragx, dragy)
                                    headerclick = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\Headerslocation.png")
                                    headerx, headery = p.center(headerclick)
                                    time.sleep(2)
                                    p.click(headerx, headery)
                                    responselocation = True
                            else:
                                errors = open(lp.accesslogpath, "a")
                                errors.write("[ %s ] Xml Intercettato\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                                errors.close()
                                responselocation = False
                                pacoloading = False
            ieri = dt.datetime.today() - dt.timedelta(g)
            mese = ieri.month
            anno = ieri.year
            ndayfrow = 8 - wday[str(dt.date(anno, mese, 1).weekday())]
            h = 0
            for j in range(2):
                if j == 1:
                    h = 113
                p.click(535 + h, 483)
                time.sleep(2)
                if dt.datetime.today().day <= g:
                    p.click(535 + h, 483)  # passo tra i calendari
                    time.sleep(3)
                giorno = dt.datetime.today() - dt.timedelta(g)
                giorno = giorno.day
                ycal = (giorno - ndayfrow) // 7 + 1
                xcal = wday[str(dt.date(anno, mese, giorno).weekday())]
                xcoord = coord[0] + h + (25 * xcal)
                ycoord = coord[1] + (20 * ycal)
                time.sleep(2)
                p.click(xcoord, ycoord)
                time.sleep(2)
            p.press("tab")
            p.press("down")
            time.sleep(1)
            p.press("down")
            time.sleep(1)
            p.press("down")
            time.sleep(1)
            p.press("tab")
            time.sleep(1)
            p.press("down")
            time.sleep(1)
            if table == "ced_paco_richiamate_fe_191":
                p.press("down")
                time.sleep(1)
            macroesigenza = True
            while macroesigenza:
                time.sleep(2)
                macroesigenza = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\macroesigenza.png")
                if macroesigenza is None:
                    p.press("tab")
                    time.sleep(5)
                    p.keyDown("ctrl")
                    p.press("down")
                    time.sleep(1)
                    passomenu = 25
                    base = 600
                    if table == "ced_paco_richiamate_fe_187":
                        for i in range(3):
                            for j in range(5):
                                p.click(523, (passomenu*j)+base)
                                time.sleep(1)
                            for z in range(5):
                                p.click(678, 700)
                                time.sleep(1)
                    else:
                        for i in range(3):
                            if i != 2:
                                for j in range(5):
                                    p.click(523, (passomenu*j)+base)
                                    time.sleep(1)
                                for z in range(5):
                                    p.click(678, 700)
                                    time.sleep(1)
                            else:
                                for x in range(4, 5):
                                    p.click(523, (passomenu*x)+base)
                                    time.sleep(1)
                                for z in range(2):
                                    p.click(678, 700)
                                    time.sleep(1)
                    p.keyUp("ctrl")
                    p.press("tab", presses=4)
                    time.sleep(2)
                    p.click(676, 711)
                    macroesigenza = False
                else:
                    macroesigenza = True
                    stuck += 1
                    if stuck == 10:
                        errors = open(lp.errorlogpath, "a")
                        errors.write("[ %s ] Sito Web Bloccato, riavvio automa %s  di 3 tentativi in corso\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), (t)))
                        errors.close()
                        t += 1
                        update_richiamate()
                        if t == 3:
                            errors = open(lp.errorlogpath, "a")
                            errors.write("[ %s ] Impossibile recuperare dati su Paco - File non scaricato o illeggible\n\r"
                                         % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                            errors.close()
                            m.mail_error("Paco - Richiamate")
                        else:
                            pass
                    else:
                        pass


def xmlextract(table):
    global t
    passox = 16
    passoxdrag = 400
    rescomplete = True
    while rescomplete:
        time.sleep(5)
        if risposta == "https://10.104.100.140/Stats/OpenAccess/NuoveAbbRip/openAccess.aspx":
            time.sleep(2)
            h = 0
            trova = True
            while trova:
                openaccess = p.locateOnScreen(r'C:\Apache\htdocs\CED_Automi\automi\image\paper.png')
                if openaccess is not None:
                    x, y = p.center(openaccess)
                    x2 = x + 8
                    time.sleep(5)
                    p.click(x2, y)
                    trova = False
            ultimares = True
            while ultimares:
                time.sleep(2)
                y2 = y + (passo*h)
                time.sleep(3)
                p.click(x2, y2)
                time.sleep(3)
                check = p.pixel(x2, y2)
                findblue = (33, 150, 243)
                if check == findblue:
                    h += 1
                    ultimares = True
                else:
                    y2 = y2 - passo
                    p.click(x2, y2)
                    ultimares = False
            time.sleep(2)
            headerclick = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\Headerslocation.png")
            headerx, headery = p.center(headerclick)
            time.sleep(2)
            p.click(headerx, headery)
            time.sleep(1)
            responselocation = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\responselocation.png")
            time.sleep(2)
            if responselocation is None:
                errors = open(lp.accesslogpath, "a")
                errors.write("[ %s ] Avvio intercettazione Xml in corso...\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                errors.close()
                responselocation = True
            else:
                buttresx, buttresy = p.center(responselocation)
                time.sleep(2)
                p.click(buttresx, buttresy)
                time.sleep(2)
                dragcheck = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\Iniziodrag.png")
                if dragcheck is None:
                    errors = open(lp.accesslogpath, "a")
                    errors.write("[ %s ] Xml in Caricamento\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    errors.close()
                    responselocation = True
                else:
                    dragx, dragy = p.center(dragcheck)
                    dragx = dragx + passox
                    time.sleep(2)
                    p.click(dragx, dragy)
                    time.sleep(2)
                    dragreal = p.dragRel(passoxdrag, 0, 2, button='left')
                    p.keyDown('ctrl')
                    p.press('c')
                    p.keyUp('ctrl')
                    xmlcontrol = pc.paste()
                    if xmlcontrol is None or xmlcontrol != lp.xmlhead:
                        errors = open(lp.errorlogpath, "a")
                        errors.write("[ %s ] Xml estratto non valido, riavvio automa %s di 3 tentativi in corso\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), (t)))
                        errors.close()
                        p.click(dragx, dragy)
                        headerclick = p.locateOnScreen(r"C:\Apache\htdocs\CED_Automi\automi\image\Headerslocation.png")
                        headerx, headery = p.center(headerclick)
                        time.sleep(2)
                        p.click(headerx, headery)
                        responselocation = True
                        t += 1
                        update_richiamate()
                        if t == 3:
                            errors = open(lp.errorlogpath, "a")
                            errors.write("[ %s ] Impossibile recuperare dati su Paco - File non scaricato o illeggible\n\r"
                                         % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                            errors.close()
                            m.mail_error("Paco - Richiamate")
                    else:
                        errors = open(lp.accesslogpath, "a")
                        errors.write("[ %s ] Xml Intercettato\n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                        errors.close()
                        time.sleep(1)
                        p.click(x2, y2)
                        time.sleep(1)
                        p.rightClick()
                        time.sleep(2)
                        p.press("down", presses=4, interval=0.2)
                        time.sleep(2)
                        p.press("right", interval=0.2)
                        time.sleep(2)
                        p.press("down", presses=3, interval=0.2)
                        time.sleep(3)
                        p.press("enter")
                        filefromcopy = pc.paste()
                        filefromcopy = filefromcopy.replace("list", "listfile")
                        open(lp.uploadpath + "\\Richiamate.xml", "w+").write(filefromcopy)
                        responselocation = False
                        rescomplete = False
                        ffq.dffrmodule(table)


def update_richiamate():
    table = ["ced_paco_richiamate_fe_187", "ced_paco_richiamate_fe_191"]
    global vpncheck
    crs = sql.createCursor()
    query = sql.cred_query("ALL")
    crs.execute(query)
    credlist = crs.fetchone()
    for i in range(0, len(table)):
        vpncheck = chc.tim_vpn("Richiamate")
        if vpncheck:
            loginpaco(credlist)
            tab.Network.requestWillBeSent = request_will_be_sent
            paco_ripetute(4, table[i])
            xmlextract(table[i])
            tab.stop()
    crs.close()



update_richiamate()
