import functions.check_connect as cc
import functions.fromfiletosql_opera as ffo
import functions.mailsender as m
import subprocess
import pyautogui as p
import sys
import time
import os
import datetime as dt

sys.path.insert(0, r"../.")

import var.links_paths as lp
import var.sql as sql

crs = sql.createCursor()
cnx = sql.cnx


def opera_conn():
    os.system("taskkill /f /t /im chrome.exe")
    keys = ["tab", "down", "space", "enter", "up", "right"]
    service = ["11", "12"]
    operalink = [lp.OperaFE, lp.OperaBO]
    tableindex = ["FE", "BO"]
    tabber = ["08", "06"]
    lines = [30, 28]
    table = {tableindex[0]: ["ced_opera_sintesi_volumi_fe_home", "ced_opera_sintesi_volumi_fe_office"],
             tableindex[1]: ["ced_opera_sintesi_volumi_bo_home", "ced_opera_sintesi_volumi_bo_office"]}
    filename = ["\ReportSintesi Volumi FE MOI.xlsx", "\ReportSintesi Volumi BO MOI.xlsx"]
    i = 0
    query = sql.cred_query("ALL")
    crs.execute(query)
    credlist = crs.fetchone()

    if dt.datetime.now().day >= 16:
        month = ["42"]
        monthyear = [str(dt.datetime.now().month - 2).zfill(2) + str(dt.datetime.now().year)[2:]]
    elif dt.datetime.now().day < 16:
        month = ["41", "31"]
        monthyear = [str(dt.datetime.now().month - 1).zfill(2) + str(dt.datetime.now().year)[2:],
                     str(dt.datetime.now().month).zfill(2) + str(dt.datetime.now().year)[2:]]
    for k in range(0, 2):
        for z in range(0, len(month)):
            for i in range(0, 2):
                keypress = ["01", "11", "21", service[i], "21", "31", "01", "11",
                            "12", "31", tabber[k], "12", "31", "04", "11", "02", "11",
                            month[z], "32", "07", "21"]
                subprocess.Popen('"%s" --ignore-certificate-errors --new-window --run-all-flash-in-allow-mode --test-type --start-maximized "%s"' % (lp.pathchrome, operalink[k]))
                p.click(cc.load(r'.\image\chrome_white_x.png', mail=0, checkload=55))
                p.click(575, 400)
                p.press("tab")
                p.typewrite(credlist[0], interval=0.1)
                p.press("tab")
                p.typewrite(credlist[1], interval=0.1)
                p.press("enter")
                if cc.load(r'.\image\opera_title.png') is not None:
                    p.click(cc.load(r'.\image\opera_title.png'))
                    for j in range(0, len(keypress)):
                        p.press(keys[int(keypress[j][:1])], presses=int(keypress[j][1:]), interval=0.5)
                        time.sleep(1)
                        if (j == 7) or (j == 11) or j == 19:
                            time.sleep(5)
                        if j == 20:
                            if k == 0:
                                if cc.load(r'.\image\opera_aperture.png') is not None:
                                    p.click(cc.load(r'.\image\opera_export.png'))
                                    p.click(cc.load(r'.\image\opera_excel1.png'))
                                    p.press("down")
                                    time.sleep(1)
                                    p.press("enter")
                                    time.sleep(30)
                                    p.click(cc.load(r'.\image\opera_exit.png'))
                            else:
                                if cc.load(r'.\image\opera_bacino.png') is not None:
                                    p.click(cc.load(r'.\image\opera_export.png'))
                                    p.click(cc.load(r'.\image\opera_excel1.png'))
                                    p.press("down")
                                    time.sleep(1)
                                    p.press("enter")
                                    time.sleep(30)
                                    p.click(cc.load(r'.\image\opera_exit.png'))
                    os.system("taskkill /f /t /im chrome.exe")
                    try:
                        ffo.cleanerOP(table[tableindex[k]][i], lines[k], lp.autodlpath + filename[k], monthyear[z])
                        os.remove(lp.autodlpath + filename[k])
                        crs.execute(sql.upddbtracker("OpeRA"))
                        cnx.commit()
                    except Exception as error:
                        errors = open(lp.errorlogpath, "a")
                        errors.write("[ %s ] Impossibile aggiornare %s sulla tabella %s - %s - %s\n\r"
                                     % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), table[tableindex[k]][i], "OpeRA", str(type(error)), str(error)))
                        errors.close()
                        m.mail_error("OpeRA")
    crs.close()
    cnx.close()


def update_opera():
    cc.tim_vpn("Opera")
    opera_conn()


update_opera()
