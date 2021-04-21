from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from string import Template
import smtplib
import datetime as dt
import subprocess
import sys
import os
import datetime as dt
import templates.aut_mail_tamplate as amt
import functions.check_connect as chc
import functions.dataframeoperations as dfo
import var.links_paths as lp
import var.sql as sql

def mail_c87():
    mdict = dict()
    ivrsmonth = dict()
    ivrsweek = dict()
    ivrstoday = dict()
    ivrsyest = dict()


    cnx = sql.cnx
    crs = sql.createCursor()

    month = dt.datetime.now().month
    year = dt.datetime.now().year
    if dt.date(year, month, 1).isoweekday() == 1:
        date = dt.date(year, month, 1).strftime("%Y-%m-%d")
    else:
        date = (dt.date(year, month, 1) - dt.timedelta(dt.date(year, month, 1).isoweekday() - 1)).strftime("%Y-%m-%d")

    table = ["ced_c87_ivr_187", "ced_c87_ivr_191"]

    querymonth = "SELECT name, number FROM ced_month_tr WHERE number = '%s'" % str(month).zfill(2)
    crs.execute(querymonth)
    mlist = crs.fetchone()
    mdict[mlist[0]] = mlist[1]

    lweek = int(dt.datetime.now().strftime("%W"))

    for i in range(len(table)):
        df = pd.read_sql("SELECT * FROM " + table[i] + " WHERE str_to_date(DATA_INT, '%d/%m/%Y') >= '"+ date + "'", cnx)
        ivrsmonth[table[i]] = dfo.oldivrmonth(df, table[i], mdict, str(year), 1, "ENNOVA")
        ivrsweek[table[i]] = dfo.oldivrweek(df, table[i], lweek, lweek, "ENNOVA")
        today = dt.datetime.now().strftime("%d/%m/%Y")
        yesterday = (dt.datetime.now() - dt.timedelta(1)).strftime("%d/%m/%Y")
        dftoday = df[df["DATA_INT"] == today]
        dfyest = df[df["DATA_INT"] == yesterday]
        ivrstoday[table[i]] = dfo.oldivrmonth(dftoday, table[i], mdict, str(year), 1, "ENNOVA")
        ivrstoday[table[i]][0]["Periodo"] = today
        ivrsyest[table[i]] = dfo.oldivrmonth(dfyest, table[i], mdict, str(year), 1, "ENNOVA")
        ivrsyest[table[i]][0]["Periodo"] = yesterday

    htmlday = amt.ivr_template_day(ivrstoday, ivrsyest, table)
    htmlweek = amt.ivr_template_wm(ivrsweek, table)
    htmlmonth = amt.ivr_template_wm(ivrsmonth, table)

    queryupdate = "SELECT date_hour FROM ced_db_update_tracker WHERE db_name = 'IVR' ORDER BY id"
    crs.execute(queryupdate)
    maildate = crs.fetchone()

    html = ("""<p>Ciao,</p>
                 <p>di seguito i dati dei survey per il servizio C87 aggiornati il """ + maildate[0][:10] + """ alle """ + maildate[0][10:] + """</p>
                 <p>Giornaliero:</p>""" + htmlday + """
                 <p>Settimanale:</p>""" + htmlweek + """
                 <p>Mensile:</p>""" + htmlmonth + """
                 <p>Buon lavoro,</p>
                 <p>Ennova Cagliari CED</p> """)


    sender = "SELECT * FROM ced_tools_credentials WHERE description = 'Mail'"
    crs.execute(sender)
    sender = crs.fetchone()

    message = MIMEMultipart("alternative", None, [MIMEText(html, 'html')])
    message['Subject'] = "IVR " + dt.datetime.now().strftime("%d/%m/%Y %H:%M")
    message['From'] = sender[0]
    message['To'] = ", ".join(lp.dl_ivr)
    receiver = lp.dl_ivr
    server = smtplib.SMTP_SSL(lp.smtpsrv, port=465)
    server.login(sender[0], sender[1])
    server.sendmail(sender[0], receiver, message.as_string())
    server.quit()

    crs.close()
    cnx.close()


def alertupd(what, where):

    cnx = sql.cnx
    crs = sql.createCursor()

    html = ("""<p>Ciao,</p>
                 <p>Ho aggiunto</p> """)
    for i in range(len(what)):
        html += "<p>" + what[i] + "</p>"
    html += ("""<p>durante l'aggiormento di """ + where + """</p>
                 <p>Buon lavoro,</p>
                 <p>Ennova Cagliari CED</p> """)


    sender = "SELECT * FROM ced_tools_credentials WHERE description = 'Mail'"
    crs.execute(sender)
    sender = crs.fetchone()

    message = MIMEMultipart("alternative", None, [MIMEText(html, 'html')])
    message['Subject'] = "IVR " + dt.datetime.now().strftime("%d/%m/%Y %H:%M")
    message['From'] = sender[0]
    message['To'] = ", ".join(lp.dl_ivr)
    receiver = lp.dl_ivr
    server = smtplib.SMTP_SSL(lp.smtpsrv, port=465)
    server.login(sender[0], sender[1])
    server.sendmail(sender[0], receiver, message.as_string())
    server.quit()

    crs.close()
    cnx.close()
