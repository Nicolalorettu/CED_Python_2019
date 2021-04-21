import xml.etree.ElementTree as ET
import xlwt
import time
import pandas as pd
import sys
import os
try:
    import functions.mailsender as m
except ModuleNotFoundError:
    import mailsender as m
import update_mail as um

sys.path.insert(0, r"../.")

import var.sql as sql
import var.links_paths as lp



def dffrmodule(table):
    errorfield = ""
    crs = sql.createCursor()
    cnx = sql.cnx
    dfcols = ['DayId', 'GroupId', 'FieldId', 'Chiamate', 'Abbattute', 'Richiamate_Soglia',
              'Richiamate_Giorno', 'Richiamate_3_Giorni', 'Richiamate_nella_settimana']
    df = pd.DataFrame(columns=dfcols)
    pathxml = lp.pathxml
    tree = ET.parse(pathxml)
    root = tree.getroot()
    crs.execute("SELECT CENTRO FROM ced_paco_mos_moi ORDER BY ID")
    end = crs.fetchall()
    #root.tag = 'listfile'
    for listfile in root:
        for item in listfile:
            if item.attrib.get('a') == 'TOTALE':
                break
            for item in item:
                for item in item:
                    for item in item:
                        for item in item:
                            if str(item.attrib.get('a')).find("ENNOVA") > -1:
                                GroupId = item.attrib.get('a')
                                chiamate = item.attrib.get('b')
                                Abbattute = item.attrib.get('c')
                                Richiamate_Soglia = item.attrib.get('e')
                                Richiamate_Giorno = item.attrib.get('g')
                                Richiamate_3_Giorni = item.attrib.get('i')
                                Richiamate_nella_settimana = item.attrib.get('m')
                                servizio = item.attrib.get('id')
                                check = True
                                i = 0
                                while check:
                                    try:
                                        if servizio.find(end[i][0]) == -1:
                                            i += 1
                                        else:
                                            FieldId = servizio[9:servizio.find(end[i][0])]
                                            check = False
                                    except IndexError:
                                        if errorfield == servizio[9:servizio.find(end[i][0])]:
                                            None
                                        else:
                                            errorfield = servizio[9:servizio.find(end[i][0])]
                                            errors = open(lp.errorlogpath, "a")
                                            errors.write("Gruppo [ %s ] non registrato nel DB\n\r" % servizio)
                                            errors.close()
                                DayId = servizio[:9].lower()
                                df = df.append(
                                    pd.Series([DayId, GroupId, FieldId, chiamate, Abbattute,
                                              Richiamate_Soglia,
                                              Richiamate_Giorno,
                                              Richiamate_3_Giorni,
                                              Richiamate_nella_settimana], index=dfcols), ignore_index=True)
# UPDATE GROUP AND SERVICE DB
    grouplist = list(df["GroupId"].unique())
    serviceslist = list(df["FieldId"].unique())

    group = pd.read_sql("SELECT * FROM ced_c87_group", sql.cnx, index_col="ID")
    services = pd.read_sql("SELECT * FROM ced_paco_services", sql.cnx, index_col="ID")
    Mesi = pd.read_sql("SELECT * FROM ced_month_tr", sql.cnx, index_col="initials")
    serviceslist = [i for i in serviceslist[0:] if i not in services.SERVICE.tolist()]
    grouplist = [i for i in grouplist[0:] if i not in group.BACINO.tolist()]

    if grouplist != []:
        errors = open(lp.errorlogpath, "a")
        errors.write("Gruppo [ %s ] non registrato nel DB\n\r" % serviceslist)
        errors.close()

    if serviceslist != []:
        errors = open(lp.errorlogpath, "a")
        errors.write("Gruppo [ %s ] non registrato nel DB\n\r" % grouplist)
        errors.close()

    # OLD UPDATE GROUPS AND SERVICES
    #change = []

    #for i in range(0, len(services.index)):
    #    change.append(int(services.index[i]))
    #services.index = change

    #change = []
    #for i in range(0, len(group.index)):
    #    change.append(int(group.index[i]))
    #group.index = change

    #for i in range(0, len(serviceslist)):
    #    query = ("INSERT INTO ced_paco_services (`ID`, `SERVICE`) VALUES ( %s, '%s')" % (str(i+max(services.index.tolist())+1), serviceslist[i]))
    #    crs.execute(query, sql.cnx)
    #    sql.cnx.commit()

    #for i in range(0, len(grouplist)):
    #    query = ("INSERT INTO ced_c87_group (`ID`, `BACINO`) VALUES ( %s, '%s')" % (str(i+max(group.index.tolist())+1), grouplist[i]))
    #    print(query)
    #    crs.execute(query, sql.cnx)
    #    sql.cnx.commit()

    #if len(grouplist) > 0:
    #    um.alertupd(grouplist, "Richiamate")

    #if len(serviceslist) > 0:
    #    um.alertupd(serviceslist, "Richiamate")

    #group = pd.read_sql("SELECT * FROM ced_c87_group", sql.cnx, index_col="ID")
    #services = pd.read_sql("SELECT * FROM ced_paco_services", sql.cnx, index_col="ID")
# END UPDATE GROUP AND SERVICE DB

# DATAFRAME PREPARATION
    for i in range(0, len(df.index)):
        try:
            df.replace(df["FieldId"][i], services.index[services.SERVICE == df["FieldId"][i]].tolist()[0], inplace=True)
        except IndexError:
            None
        try:
            df.replace(df["GroupId"][i], group.index[group.BACINO == df["GroupId"][i]].tolist()[0], inplace=True)
        except IndexError:
            None
        try:
            df.replace(df["DayId"][i], (df["DayId"][i][0:3] + Mesi.number[Mesi.index == df["DayId"][i][3:6].capitalize()].tolist()[0] + df["DayId"][i][6:]), inplace=True)
        except IndexError:
            None

    id = []
    for i in range(0, len(df.index)):
        id.append(str(df["DayId"][i]) + str(df["GroupId"][i]) + str(df["FieldId"][i]))
    df.insert(loc=0, column="ID", value=id)

# UPDATE TABLE SINTESI
    df.to_csv("c:/ProgramData/MySQL/MySQL Server 8.0/Uploads/cleanpacodf.csv", sep=";", index=False)

    query = sql.repwithcsv("/ProgramData/MySQL/MySQL Server 8.0/Uploads/cleanpacodf.csv", table, "1")
    crs.execute(query, sql.cnx)
    cnx.commit()
    if os.path.exists(lp.pathxml):
        os.remove(lp.pathxml)
        access = open(lp.accesslogpath, "a")
        access.write("Dato caricato sulla tabella [ %s ].\n\r" % (table))
        access.close()
    else:
        errors = open(lp.errorlogpath, "a")
        errors.write("Impossibile trovare %s sulla cartella indicata - File non scaricato o illeggible \n\r"
                     % (lp.pathxml))
        errors.close()
        m.mail_error("Paco - Richiamate")
