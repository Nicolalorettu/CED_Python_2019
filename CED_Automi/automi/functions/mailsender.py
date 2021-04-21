import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

sys.path.insert(0, r"../.")

import var.links_paths as lp
import var.sql as sql

crs = sql.createCursor()


def mail_error(updateproc):
    query = sql.cred_query("Mail")
    crs.execute(query)
    credlist = crs.fetchone()
    body = ("""
            Non è stato possibile effettuare l'aggiornamento di %s.
            Si prega di verificare ed aggiornare manualmente.

            Messaggio inviato a:
            %s
           """ % (updateproc, lp.dl_errors))
    message = MIMEMultipart()
    message.attach(MIMEText(body, "plain"))
    message["Subject"] = "Errore Update SQL"
    message["From"] = credlist[0]
    message["To"] = ", ".join(lp.dl_errors)
    filename = "errorlog.txt"
    attachment = open(lp.errorlogpath, "r")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    message.attach(part)
    server = smtplib.SMTP_SSL(lp.smtpsrv, port=465)
    server.login(credlist[0], credlist[1])
    server.sendmail(credlist[0], lp.dl_errors, message.as_string())
    server.quit()


def mail_PACO(filenames):
    query = sql.cred_query("Mail")
    crs.execute(query)
    credlist = crs.fetchone()
    body = ("""
    In allegato l'estrazione di PACO - TMC - Kpi Aso

    Buon lavoro,
    ENNOVA CED
           """)
    message = MIMEMultipart()
    message.attach(MIMEText(body, "plain"))
    message["Subject"] = "Automa PACO - TMC - Kpi Aso"
    message["From"] = credlist[0]
    message["To"] = ", ".join(lp.dl_errors)
    for filename in filenames:
        attachment = open(lp.autodlpath + filename, "r")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        message.attach(part)
    server = smtplib.SMTP_SSL(lp.smtpsrv, port=465)
    server.login(credlist[0], credlist[1])
    server.sendmail(credlist[0], lp.dl_errors, message.as_string())
    server.quit()
