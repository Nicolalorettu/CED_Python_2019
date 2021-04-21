import subprocess
import time
import os
import pyautogui as p
import functions.mailsender as m
import sys
import datetime as dt

sys.path.insert(0, r"../.")

import var.links_paths as lp
import var.sql as sql

crs = sql.createCursor()


def tim_vpn(updateproc):
    crs = sql.createCursor()
    i = 0
    query = sql.cred_query("ALL")
    crs.execute(query)
    credlist = crs.fetchone()
    vpncheck = True
    googlecode = False
    googlecode = os.system("ping -n 1 " + lp.Googleurl)
    if googlecode != 0:
        try:
            os.system("taskkill /f /t /im ipseca.exe")
            googlecode = os.system("ping -n 1 " + lp.Googleurl)
        except:
            errors = open(lp.errorlogpath, "a")
            errors.write("[ %s ] Impossibile collegarsi a %s - Host non raggiungibile o Rete KO \n\r"
                         % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), lp.Googleurl))
            errors.close()
    while googlecode is 0:
        subprocess.Popen(lp.pathvpn)
        time.sleep(3)
        p.press("right", presses=3, interval=0.5)
        p.press("alt")
        p.press("down")
        p.press("enter")
        time.sleep(1)
        p.typewrite(credlist[0], interval=0.1)
        time.sleep(2)
        p.press("tab")
        p.typewrite(credlist[1], interval=0.1)
        time.sleep(2)
        p.press("enter")
        time.sleep(10)
        googlecode = os.system("ping -n 1 " + lp.Googleurl)
        print(googlecode)
        if googlecode == 0:
            os.system("taskkill /f /t /im " + lp.prcvpn)
            time.sleep(10)
            i = i + 1
        if i > 4:
            googlecode = False
            vpncheck = False
            errors = open(lp.errorlogpath, "a")
            errors.write("[ %s ] Impossibile collegarsi a %s - Host non raggiungibile o VPN KO \n\r"
                         % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), updateproc))
            errors.close()
            m.mail_error(updateproc)
    return vpncheck


def free_net(updateproc):
    i = 0
    googlecode = os.system("ping -n 1 " + lp.Googleurl)
    time.sleep(1)
    timcode = os.system("ping -n 1 " + lp.Timurl)
    time.sleep(1)
    if (googlecode != 0 and timcode != 0):
        try:
            os.system("taskkill /f /t /im ipseca.exe")
            time.sleep(1)
        except:
            None
    googlecode = os.system("ping -n 1 " + lp.Googleurl)
    time.sleep(1)
    timcode = os.system("ping -n 1 " + lp.Timurl)
    time.sleep(1)
    if (googlecode != 0 and timcode != 0):
        errors = open(lp.errorlogpath, "a")
        errors.write("[ %s ] Impossibile collegarsi a %s - Server non raggiungibile o Rete KO \n\r"
                     % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), updateproc))
        errors.close()
        m.mail_error(updateproc)
    else:
        neton = True
    return neton


def load(path, mail=1, checkload=0):
    loading = True
    while loading:
        if p.locateCenterOnScreen(path) is not None:
            x, y = p.locateCenterOnScreen(path)
            loading = False
            return x, y
        else:
            checkload += 1
        time.sleep(1)
        if checkload == 60:
            if mail == 1:
                os.system("taskkill /f /t /im ipseca.exe")
                errors = open(lp.errorlogpath, "a")
                errors.write("[ %s ] Impossibile recuperare dati su Paco/Ibia - %s non trovato\n\r"
                             % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), path))
                errors.close()
                m.mail_error("")
            else:
                return None
            return None
