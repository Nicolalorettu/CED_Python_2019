import xml.etree.ElementTree as ET
import datetime as dt
import xlwt
import time
import pandas as pd
import sys
import os
import ast
import automi.var.sql as sql
import automi.var.links_paths as lp

crs = sql.createCursor()
cnx = sql.cnx


def ASCupdIBIA(data):
    # FROM FILE TO DATAFRAME TABELLA ASC E AGGIORNAMENTO DB ASC
    dfgruppi = pd.DataFrame()
    data = data.replace("\n", "")

    data = data[data.find('"Members"')+10:data.find(',"More"')]

    data = data.replace("true", "True")
    data = data.replace("false", "False")
    data = data.replace("ENNOVA_CA ATPAY", "ENNOVA_CA_ATPAY")
    dictgruppi = ast.literal_eval(data)

    for i in range(0, len(dictgruppi)):
        dfgruppi = pd.concat([dfgruppi, pd.DataFrame(dictgruppi[i], index=[0])], axis=0, sort=False, ignore_index=True)

    for i in range (0, len(dfgruppi.index)):
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

    ASClist = list(dfgruppi["Id"].unique())
    dbASC = pd.read_sql("SELECT * FROM ced_c87_ASC", cnx)
    ASClist = [i for i in ASClist[1:] if i not in dbASC.ASC.tolist()]

    for i in range(0, len(ASClist)):
        query = ("INSERT INTO ced_c87_ASC (`ID`, `ASC`) VALUES ( %s, '%s')" % (str(i+int(max(dbASC.ID.tolist()))+1), ASClist[i]))
        crs.execute(query, cnx)
        cnx.commit()
    # END FROM FILE TO DATAFRAME TABELLA ASC E AGGIORNAMENTO DB ASC
    # AGGIORNAMENTO DB GRUPPI CON ASC
    dbASC = pd.read_sql("SELECT * FROM ced_c87_ASC", cnx)
    dbgroup = pd.read_sql("SELECT * FROM ced_c87_group", cnx)

    for i in range (0, len(dfgruppi.index)):
        if dfgruppi["Lv"][i] == 2 or dfgruppi["Lv"][i] == 1:
            dfgruppi.drop(i, inplace=True)

    for i in range(0, max(dfgruppi.index)):
        for j in range(0, len(dbASC["ASC"])):
            try:
                if dbASC["ASC"][j] == dfgruppi["Id"][i]:
                    dfgruppi["Lv"][i] = dbASC["ID"][j]
            except KeyError:
                None

    groupnotlist = [i for i in dfgruppi["Tx"] if i in dbgroup.BACINO.tolist()]
    for i in range(0, max(dfgruppi.index)):
        for j in range(0, len(groupnotlist)):
            try:
                if dfgruppi["Tx"][i] == groupnotlist[j]:
                    dfgruppi.drop(i, inplace=True)
            except KeyError:
                None

    j = 1
    for i in range(0, max(dfgruppi.index)):
        try:
            query = ("INSERT INTO ced_c87_group (ID, BACINO, Keys_ASC_id) VALUES (%s, '%s', %s)" % (str(int(max(dbgroup["ID"]))+j), dfgruppi["Tx"][i], str(dfgruppi["Lv"][i]).zfill(2)))
            crs.execute(query, cnx)
            cnx.commit()
            j += 1
        except KeyError:
            None
    # END   AGGIORNAMENTO DB GRUPPI CON ASC
    return grouplen


def cleanerIBIA(data):
    filename = open(lp.automidlpath + "\ibiatable.html", 'w')
    filename.write(data)
    filename.close()
# IBIA converter from string to df for DB AND UPDATE DB
    service = ""
    columns = list()
    dictgruppi = list()
    dfgruppi = pd.DataFrame()
    dfservizi = pd.DataFrame()
    servizilist = list()
    rowslist = list()
    columns = list()

    # FROM FILE TO DATAFRAME TABELLA VALORI
    dftable = pd.read_html(filename, keep_default_na=False)
    dftable = dftable[0]
    for i in range(0, len(dftable.columns)):
        columns.append(dftable.loc[dftable.index[0], i])
    dftable.drop(0, inplace=True)
    dftable.columns = columns
    # END FROM FILE TO DATAFRAME TABELLA VALORI

    # AGGIORNAMENTO DB RIGHE
    rowslist = dftable["DIAGNOSI FE MACRODES"].unique()
    dbrows = pd.read_sql("SELECT * FROM ced_c87_ibia_rows", cnx, index_col="ID")
    rowslist = [i for i in rowslist[1:] if i not in dbrows.ROW.tolist()]
    for i in range(0, len(rowslist)):
        query = ("INSERT INTO ced_c87_ibia_rows (`ID`, `ROW`) VALUES ( %s, '%s')"
                 % (str(i+int(max(dbrows.index.tolist()))+1), rowslist[i]))
        crs.execute(query, cnx)
        cnx.commit()
    # END AGGIORNAMENTO DB RIGHE

    # INTEGRATE DATAFRAME WITH FOREIGN KEYS AND UPDATE DB
    dbgroup = pd.read_sql("SELECT * FROM ced_c87_group", cnx)

    with open(filename, "r") as myfile:
        data=myfile.read().replace("\n", "")

    stringinit = data.find(".[TEAM SEGN].")+20
    partial = data[stringinit:]
    group = data[stringinit:stringinit+partial.find(']", "')]
    stringinit = data.find(" Data Segnalazione Gerarchia: ")+30
    partial = data[stringinit:]
    year = dt.datetime.now().strftime("%Y")
    date = data[stringinit:stringinit+partial.find(year)+5]
    stringinit = data.find(".[TIPOLOGIA CLIENTE].")+28
    partial = data[stringinit:]
    if partial.find('tion]", "') == -1:
        table = data[stringinit:stringinit+partial.find('].')]
    else:
        table = data[stringinit+1:stringinit+partial.find('tion]", "')+4]
        service = "ADSL"
    stringinit = data.find("u0026[DATI].")+19
    if data.find("u0026[DATI].") == -1:
        stringinit = data.find("u0026[FONIA].")+20
    print(stringinit)
    partial = data[stringinit:]
    if service == "":
        service = data[stringinit:stringinit+partial.find(']", "')]

    print(group)
    print(date)
    print(table)
    print(service)

    day = str(int(date[:2])).zfill(2)
    month = date[date.find(" ")+1:date.find(year)-1]
    query = ("SELECT number FROM ced_month_tr WHERE name = '%s'" % month)
    crs.execute(query, cnx)
    month = crs.fetchone()
    date = day + "/" + month[0] + "/" + year

    query = ("SELECT ID FROM ced_c87_group WHERE BACINO = '%s'" % group)

    crs.execute(query, cnx)
    bacino = crs.fetchone()

    query = ("SELECT ID FROM ced_c87_ibia_service WHERE SERVICE = '%s'" % service)

    crs.execute(query, cnx)
    service = crs.fetchone()
    for i in range(1, len(dftable[columns[0]])+1):
        query = ("SELECT ID FROM ced_c87_ibia_rows WHERE ced_c87_ibia_rows.ROW = '%s'" % dftable[columns[0]][i])
        crs.execute(query, cnx)
        row = crs.fetchone()
        dftable[columns[0]][i] = row[0]

    dftable.insert(loc=0, column="Keys_SERVIZIO", value=service[0])
    dftable.insert(loc=0, column="Keys_BACINO", value=bacino[0])
    dftable.insert(loc=0, column="Date", value=date)
    dftable.insert(loc=0, column="ID", value=date+"-"+bacino[0]+"-"+service[0]+"-"+dftable[columns[0]])

    dftable.drop(["% TK Rip 1gg", "% TK Verifiche Negative"], axis=1, inplace=True)
    columns = dftable.columns
    for i in range(0, len(dftable.columns)):
        for j in range(1, len(dftable.index)+1):
            if dftable[columns[i]][j] == "":
                dftable[columns[i]][j] = 0
    # END INTEGRATE DATAFRAME WITH FOREIGN KEYS AND UPDATE DB

    if table == "BUSINESS":
        table = "ced_c87_ibia_business"
    elif table == "RESIDENZIALE" or table == "Aggregation":
        table = "ced_c87_ibia_residential"

    endfilename = "cleanIBIA.csv"
    dftable.to_csv(lp.uploadpath + endfilename, sep=";", index=False)
    query = sql.repwithcsv(lp.uploadpath + endfilename, table, "1")
    crs.execute(query)
    cnx.commit()


# END CLEANING AND UPDATE DB

crs.close()
cnx.close()
