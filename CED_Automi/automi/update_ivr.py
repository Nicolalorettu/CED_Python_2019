import functions.check_connect as cc
import functions.mailsender as m
import subprocess
import pyautogui as p
import sys
import time
import os
import datetime as dt
import functions.fromfiletosql_ivr as ffq

sys.path.insert(0, r"../.")

import var.links_paths as lp
import var.sql as sql

crs = sql.createCursor()
cnx = sql.cnx


def updmultiday(i):
    coord = [760, 540]
    # scarica un numero di giorni pari ad "i"
    oldestdate = dt.datetime.today() - dt.timedelta(i)
    month = oldestdate.month
    year = oldestdate.year
    ndayfrow = 8 - dt.date(year, month, 1).isoweekday()

    if dt.datetime.today().day <= i:
        p.click(750, 480)
        time.sleep(2)

    datefrom = dt.datetime.today() - dt.timedelta(i)
    datefrom = datefrom.day

    ycal = (datefrom - ndayfrow) // 7 + 1
    xcal = dt.date(year, month, datefrom).isoweekday()
    xcoord = coord[0] + (28 * xcal)
    ycoord = coord[1] + (22 * ycal)
    time.sleep(2)
    p.click(xcoord, ycoord)


def ivr_conn(nday):
    os.system("taskkill /f /t /im chrome.exe")
    keys = ["tab", "down", "space", "enter", "up", "right"]
    service = ["11", "12"]
    pacolink = (lp.PACOlink)
    table = ["ced_c87_ivr_187", "ced_c87_ivr_191"]
    filename = "\AssistenzaTecnica IVR.xls"
    query = sql.cred_query("ALL")
    crs.execute(query)
    credlist = crs.fetchone()

    for i in range(0, len(table)):
        keypress = [service[i], "31", "02", "11", "01", "21"]
        subprocess.Popen('"%s" --ignore-certificate-errors --new-window --run-all-flash-in-allow-mode --test-type --args --start-maximized "%s"' % (lp.pathchrome, pacolink))
        time.sleep(5)
        p.click(575, 400)
        time.sleep(1)
        p.press("tab")
        time.sleep(0.5)
        p.press("enter")
        p.click(cc.load(r'.\image\paco_ok.png'))
        if cc.load(r'.\image\paco_tim.png') is not None:
            p.press("tab", presses=2, interval=0.5)
            p.typewrite(credlist[0], interval=0.1)
            p.press("tab")
            p.typewrite(credlist[1], interval=0.1)
            p.press("enter")
            p.click(cc.load(r'.\image\paco_oatools.png'))
            for j in range(0, len(keypress)):
                p.press(keys[int(keypress[j][:1])], presses=int(keypress[j][1:]), interval=0.5)
                time.sleep(1)
                if j == 1:
                    if cc.load(r'.\image\paco_survey.png') is not None:
                        updmultiday(nday)
            time.sleep(5 * nday)
            p.click(1145, 715, button="right")
            time.sleep(2)
            p.click(1245, 775)
            time.sleep(5)
            p.typewrite(lp.autodlpath + filename)
            p.press("enter", presses=3, interval=0.5)
            time.sleep(20 * nday)
            os.system("taskkill /f /t /im chrome.exe")
            time.sleep(5)

        try:
            ffq.convert_xml(lp.autodlpath + filename, table[i])
            crs.execute(sql.upddbtracker("IVR"))
            cnx.commit()
        except:
            try:
                errors = open(lp.errorlogpath, "a")
                errors.write("[ %s ] Impossibile aggiornare %s sulla tabella %s - File non scaricato o illeggible \n\r"
                             % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), table[i], "IVR"))
                errors.close()
                m.mail_error("IVR")
            except TimeoutError:
                None
    crs.close()
    cnx.close()


def update_ivr():
    cc.tim_vpn("IVR")
    ivr_conn(1)


update_ivr()
