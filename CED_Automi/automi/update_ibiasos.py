import functions.mailsender as m
import subprocess
import pyautogui as p
import sys
import time
import os
import datetime as dt
import functions.fromfiletosql_ibiasos as ffq
import functions.check_connect as cc

sys.path.insert(0, r"../.")

import var.links_paths as lp
import var.sql as sql

crs = sql.createCursor()
cnx = sql.cnx


def start_connection(type):
    if type == "manual":
        anno = "2018"
        mese = input("Che Mese?")
        dal = input("Dal giorno?")
        al = input("al giorno?")
    else:
        anno = 0
        mese = 0
        dal = 0
        al = 0
    crs = sql.createCursor()
    query = sql.cred_query("ALL")
    crs.execute(query)
    credlist = crs.fetchone()
    vpncheck = cc.tim_vpn("Ibia")
    if vpncheck:
        loginIbia(credlist, type, anno, mese, dal, al)


def loginIbia(credlist, type, anno, mese, dal, al):
    os.system("taskkill /f /t /im iexplore.exe")
    subprocess.Popen('"%s" "%s"' % (lp.pathIE, lp.IbiaSOS))
    time.sleep(5)
    p.press("alt")
    time.sleep(0.5)
    p.press("space")
    time.sleep(0.5)
    p.press("n")
    try:
        p.click(cc.load(r'.\image\explorer_certificate.png', mail=0))
    except:
        None
    if cc.load(".\image\explorer_user.png") is not None:
        if type == "auto":
            extractfile_auto(credlist)
        elif type == "manual":
            extractfile_manual(credlist, anno, mese, dal, al)
    else:
        print("x")
        #invio mail


def extractfile_auto(credlist):
    anno = str((dt.datetime.now() - dt.timedelta(1)).year)
    nmese = str((dt.datetime.now() - dt.timedelta(1)).month).zfill(2)
    crs = sql.createCursor()
    query = sql.month_query(nmese)
    crs.execute(query)
    mese = crs.fetchone()[0]
    giorno = str((dt.datetime.now() - dt.timedelta(1)).day)
    time.sleep(10)
    p.typewrite("PARTNERS\\" + credlist[0], interval=0.1)
    time.sleep(0.2)
    p.press("tab")
    time.sleep(0.2)
    p.typewrite(credlist[1], interval=0.1)
    time.sleep(0.2)
    p.press("enter")
    if cc.load(".\image\explorer_grey.png") is not None:
        p.press("tab", presses=34, interval=0.1)
        p.press("enter")
        x, y = cc.load(".\image\explorer_selectall.png")
        p.doubleClick(x, y, button='left', interval=0.2)
        x, y = cc.load(".\image\explorer_filtro.png")
        p.click(x, y)
        p.typewrite(str(giorno) + " " + mese + " " + anno, interval=0.1)
        p.press("enter")
        p.click(cc.load(".\image\explorer_check.png"))
        p.click(cc.load(".\image\explorer_ibiasos_ok.png"))
        p.click(cc.load(".\image\explorer_ibiasos_ok.png"))
        x, y = cc.load(".\image\explorer_ibiasos_lav.png")
        p.click(x+50, y+20, button="right")
        time.sleep(1)
        p.press("down", presses=3, interval=0.2)
        p.press("right")
        p.click(cc.load(".\image\explorer_ibiasos_dettaglio.png"))
        try:
            p.click(cc.load(r'.\image\explorer_certificate.png', mail=0))
        except:
            None
        p.click(cc.load(".\image\explorer_ibiasos_export.png"))
        p.click(cc.load(".\image\explorer_save.png"))
        p.keyDown("alt")
        p.press("f4")
        p.keyUp("alt")
        print(nmese)
        try:
            ffq.convert_excel(anno, nmese, giorno)
            crs.execute(sql.upddbtracker("IBIA SOS EASY"))
            cnx.commit()
        except:
            try:
                errors = open(lp.errorlogpath, "a")
                errors.write("[ %s ] Impossibile aggiornare %s sulla tabella %s - File non scaricato o illeggible \n\r"
                             % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "ced_sos_easy_ibia", "IBIA SOS EASY"))
                errors.close()
                m.mail_error("IBIA SOS EASY")
            except TimeoutError:
                None
    crs.close()
    cnx.close()


def extractfile_manual(credlist, anno, mese, dal, al):
    time.sleep(10)
    p.typewrite("PARTNERS\\" + credlist[0], interval=0.1)
    time.sleep(0.2)
    p.press("tab")
    time.sleep(0.2)
    p.typewrite(credlist[1], interval=0.1)
    time.sleep(0.2)
    p.press("enter")
    for i in range(int(dal), int(al)+1):
        z = 5
        if i > int(dal):
            z = 0
        if cc.load(".\image\explorer_grey.png") is not None:
            p.press("tab", presses=29+z, interval=0.1)
            p.press("enter")
            x, y = cc.load(".\image\explorer_selectall.png")
            p.doubleClick(x, y, button='left', interval=0.2)
            x, y = cc.load(".\image\explorer_filtro.png")
            p.click(x, y)
            p.typewrite(str(i) + " " + mese + " " + anno, interval=0.1)
            p.press("enter")
            p.click(cc.load(".\image\explorer_check.png"))
            p.click(cc.load(".\image\explorer_ibiasos_ok.png"))
            p.click(cc.load(".\image\explorer_ibiasos_ok.png"))
            x, y = cc.load(".\image\explorer_ibiasos_lav.png")
            p.click(x+50, y+20, button="right")
            time.sleep(1)
            p.press("down", presses=3, interval=0.2)
            p.press("right")
            p.click(cc.load(".\image\explorer_ibiasos_dettaglio.png"))
            try:
                p.click(cc.load(r'.\image\explorer_certificate.png', checkload=45, mail=0))
            except:
                None
            p.click(cc.load(".\image\explorer_ibiasos_export.png"))
            p.click(cc.load(".\image\explorer_save.png"))
            p.keyDown("alt")
            p.press("f4")
            p.keyUp("alt")
            crs = sql.createCursor()
            query = sql.month_query_name(mese)
            crs.execute(query)
            nmese = crs.fetchone()[1].zfill(2)
            giorno = str(i)
            time.sleep(5)
            try:
                ffq.convert_excel(anno, nmese, giorno)
                crs.execute(sql.upddbtracker("IBIA SOS EASY"))
                cnx.commit()
            except:
                try:
                    errors = open(lp.errorlogpath, "a")
                    errors.write("[ %s ] Impossibile aggiornare %s sulla tabella %s - File non scaricato o illeggible \n\r"
                                 % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "ced_sos_easy_ibia", "IBIA SOS EASY"))
                    errors.close()
                    m.mail_error("IBIA SOS EASY")
                except TimeoutError:
                    None

start_connection("manual")
crs.close()
cnx.close()
