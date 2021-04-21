import imaplib
import email
import os
import zipfile
import pandas as pd
import sys
import string
import shutil as sh
import datetime as dt

sys.path.insert(0, r"../.")

import var.sql as sql
import var.links_paths as lp

svdir = r'C:\Apache\htdocs\CED_Automi\automi\downloads'

user = "nicola.lorettu@ennovagroup.it"
password = "nopofi16"


def lastmail():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(user, password)
    mail.select("Inbox")
    subject = "dati di avanzamento da sms online"
    typ, msgs = mail.search(None, '(SUBJECT "%s")' % subject)
    msgs = msgs[0].split()
    resp, data = mail.fetch(msgs[len(msgs)-1], '(RFC822)')
    email_body = data[0][1].decode('utf-8')
    m = email.message_from_string(email_body)
    for part in m.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        filename = part.get_filename()
        if filename is not None:
            sv_path = os.path.join(svdir, filename)
            if filename.endswith(".zip"):
                fp = open(sv_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
                atpaysms = open(lp.atpaylogpath, "a")
                atpaysms.write(" [ %s ] File zip catturato e depositato in: %s \n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), svdir))
                atpaysms.close()

    mail.close()
    mail.logout()


def findpassword():
    keyp = ['psw', 'La psw', 'La passwd', 'La pswrd']
    body = ""
    x = 0
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(user, password)
    mail.select("Inbox")
    subject = "dati di avanzamento da sms online"
    typ, msgs = mail.search(None, '(SUBJECT "%s")' % subject)
    msgs = msgs[0].split()
    for findmail in msgs:
        resp, data = mail.fetch(findmail, '(RFC822)')
        email_body = data[0][1].decode('utf-8')
        m = email.message_from_string(email_body)
        if m.is_multipart():
            for part in m.walk():
                ctype = part.get_content_type()
                cdispo = str(part.get('Content-Disposition'))
                if ctype == 'text/plain' and 'attachment' not in cdispo:
                    body = str(part.get_payload(decode=True))
                    for pas in range(0, len(keyp)):
                        check = body.find(keyp[pas])
                        if check != -1:
                            leftswipe = (body[check:check+50])
                            temp = leftswipe.replace('La passwd \\xc3\\xa8: ', '')
                            temp2 = temp.replace('*', '')
                            rightswipe = temp2.replace('\\r\\n', '')
                            finalswipe = rightswipe.replace('\\', '')
        else:
            body = str(m.get_payload(decode=True))
            for pas in range(0, len(keyp)):
                check = body.find(keyp[pas])
                if check != -1:
                    leftswipe = (body[check:check+50])
                    temp = leftswipe.replace('La passwd \\xc3\\xa8: ', '')
                    temp2 = temp.replace('*', '')
                    rightswipe = temp2.replace('\\r\\n', '')
                    finalswipe = rightswipe.replace('\\', '')

    mail.close()
    mail.logout()
    return finalswipe


def unzipfile(zippass):
    for filec in os.listdir(svdir):
        if filec.endswith(".zip"):
            pathcomp = os.path.join(svdir, filec)
            with zipfile.ZipFile(pathcomp, "r") as zip_ref:
                zip_ref.extractall(svdir, pwd=bytes(zippass, 'utf-8'))
            os.remove(pathcomp)
            atpaysms = open(lp.atpaylogpath, "a")
            atpaysms.write(" [ %s ] Rimmosso %s, estrazione effettuata in %s \n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), filec, svdir))
            atpaysms.close()


def mountdb():
    cnx = sql.cnx
    crs = sql.createCursor()
    smspd = dict()
    id = []
    databasetables = ["ced_sos_easy_sms_risposte", "ced_sos_easy_sms_fake", "ced_sos_easy_sms_non3xx"]
    responseok = ["ok", "OK", "Ok", "O k", "o K", "oK", "O K", "o k", "0k", "problema risolto", "SI ", "okay", "OKAY", "Okay", "004F006B", "D83DDC4D", "ç", "ç", "ç", "ç", "ç", "ç", "ç", "ç", "ç"]
    #responseko = ["ko", "KO", "Ko", "kO", "k o", "K O", "K o", "k O", "No non", "K0", "NO", "201C006B0030201D", "004E004F004E00200048004F0020005200490053004F004C0054004F", "No", "Non", "Lo", "Io", "K9", "persiste", "Persiste", "Ladri", "ladri", "devo pagare", "risolto da solo", "non è stato risolto"]
    esito = []
    xlsxdir = svdir + "\Ennova"
    for filexlsx in os.listdir(xlsxdir):
        if filexlsx.endswith(".xlsx"):
            list_file = os.path.join(xlsxdir, filexlsx)
            if filexlsx.find('Risposte') != -1:
                # FILE 1 SHEET 0
                smspd[filexlsx] = pd.read_excel(list_file, sheet_name=0)
                # CLEANING EXCEL FILE I SHEET 0
                smspd[filexlsx] = smspd[filexlsx].drop(['[$Template SMS].[Mittente]', '[$Template SMS].[Stato Template]', '[$Template SMS].[Creato Il]',
                                                        '[$Template SMS].[Testo SMS]', '[$Template SMS].[GRUPPO TEMPLATE]', 'DATA RISPOSTA dal CLIENTE'], axis=1)
                for i in range(0, len(smspd[filexlsx].index)):
                    id.append(str(smspd[filexlsx]["[$Dettaglio Risposte SMS].[Id Invio]"][i]) + str(smspd[filexlsx]["[$Dettaglio Risposte SMS].[Id Risposta]"][i]))
                smspd[filexlsx].insert(loc=0, column="ID", value=id)
                smspd[filexlsx]['Primo_SMS'] = "NO"
                for j in range(0, len(smspd[filexlsx].index)):
                    smspd[filexlsx]["[$Dettaglio Risposte SMS].[Testo Sms Risposta]"][j].strip()
                    table = str.maketrans({key: None for key in string.punctuation})
                    smspd[filexlsx]["[$Dettaglio Risposte SMS].[Testo Sms Risposta]"][j] = (
                                smspd[filexlsx]["[$Dettaglio Risposte SMS].[Testo Sms Risposta]"][j].translate(table))
                    check = True
                    k = 0
                    while check:
                        try:
                            #if smspd[filexlsx]["[$Dettaglio Risposte SMS].[Testo Sms Risposta]"][j].find(responseko[k]) != -1:
                            #    esito.append("KO")
                            #    check = False
                            if smspd[filexlsx]["[$Dettaglio Risposte SMS].[Testo Sms Risposta]"][j].find(responseok[k]) != -1:
                                esito.append("OK")
                                check = False
                        except IndexError:
                            #esito.append("N.V.")
                            esito.append("KO")
                            check = False
                        k += 1
                smspd[filexlsx]['Esito'] = esito
                # FILE 1 SHEET 1
                df = pd.read_excel(list_file, sheet_name=1)
                # CLEANING EXCEL FILE I SHEET 1
                df = df.drop(['[$Template SMS].[Mittente]', '[$Template SMS].[Stato Template]', '[$Template SMS].[Creato Il]',
                              '[$Template SMS].[Testo SMS]', '[$Template SMS].[GRUPPO TEMPLATE]', 'DATA RISPOSTA dal CLIENTE'], axis=1)
                df['Primo_SMS'] = "SI"
            else:
                if filexlsx.find('fake') != -1:
                    id = []
                    smspd[filexlsx] = pd.read_excel(list_file)
                    # CLEANING EXCEL FILE II
                    smspd[filexlsx] = smspd[filexlsx].drop(['[$Dettagli].[Priorita]', '[$Dim Template SMS Invii].[Eliminato Da]',
                                                            '[$Dim Template SMS Invii].[Eliminato Il]', '[$Dim Template SMS Invii].[Blocco Notturno]', '[$Dim Template SMS Invii].[DESCR TEMPLATE]',
                                                            '[$Dim Template SMS Invii].[GRUPPO TEMPLATE]', '[$Dim Template SMS Invii].[Testo SMS]', 'Data Invio'], axis=1)
                    for i in range(0, len(smspd[filexlsx].index)):
                        id.append(str(smspd[filexlsx]["[$Dettagli].[Id Invio]"][i]))
                    smspd[filexlsx].insert(loc=0, column="ID", value=id)
                else:
                    id = []
                    smspd[filexlsx] = pd.read_excel(list_file)
                    # CLEANING EXCEL FILE III
                    smspd[filexlsx] = smspd[filexlsx].drop(['VERIFICA', 'WR', 'STATO WR', 'USCITA TECNICO', 'FLG SI NO', 'CENTRALINO', 'MODALITA PAGAMENTO'], axis=1)
                    smspd[filexlsx]['DATA APP'] = smspd[filexlsx]['DATA APP'].str.replace('.', ':')
                    smspd[filexlsx]['DATA LAV'] = smspd[filexlsx]['DATA LAV'].str.replace('.', ':')
                    smspd[filexlsx]['DESCRIZIONE'] = smspd[filexlsx]['DESCRIZIONE'].str.replace("'", '')
                    for i in range(0, len(smspd[filexlsx].index)):
                        id.append(str(smspd[filexlsx]["ETT"][i]) + str(smspd[filexlsx]["DATA ETT"][i]).replace('-', '').replace(' 00:00:00', ''))
                    smspd[filexlsx].insert(loc=0, column="ID", value=id)
            columnfile = list(smspd[filexlsx].columns)
            for x in range(0, len(columnfile)):
                control = columnfile[x].find('.')
                if control != -1:
                    columnfile[x] = columnfile[x][control:control+30]
                    columnfile[x] = columnfile[x].replace('.[', '')
                    columnfile[x] = columnfile[x].replace(']', '')
            smspd[filexlsx].columns = columnfile
            smspd[filexlsx] = smspd[filexlsx].fillna(0)
            smspd[filexlsx].columns = smspd[filexlsx].columns.str.replace('BACINO BO', 'BACINO_BOF')
            smspd[filexlsx].columns = smspd[filexlsx].columns.str.replace(' ', '_')
            smspd[filexlsx].columns = smspd[filexlsx].columns.str.replace('Id', 'Cod')
            smspd[filexlsx].columns = smspd[filexlsx].columns.str.replace('check', 'Chk')
        if filexlsx == "en_2018_SMS Risposte.xlsx":
            table = databasetables[0]
        elif filexlsx == "EN_SMS_fake.xlsx":
            table = databasetables[1]
        else:
            table = databasetables[2]
    # Insert into DB
        for i in range(0, len(smspd[filexlsx].index)):
            query = sql.insert_query(table, smspd[filexlsx], i)
            crs.execute(query, cnx)
            cnx.commit()
        if filexlsx == "en_2018_SMS Risposte.xlsx":
            for j in range(0, len(df.index)):
                query2 = sql.update_SMS(table, (str(df['[$Dettaglio Risposte SMS].[Id Risposta]'][j])))
                crs.execute(query2, cnx)
                cnx.commit()
    crs.close()
    sh.rmtree(xlsxdir)
    atpaysms = open(lp.atpaylogpath, "a")
    atpaysms.write(" [ %s ] Rimossi tutti i file da %s, le tabelle: %s  sono state aggiornate. \n\r" % (dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), xlsxdir, (databasetables)))
    atpaysms.close()


lastmail()
zip_pas = findpassword()
unzipfile(zip_pas)
mountdb()
