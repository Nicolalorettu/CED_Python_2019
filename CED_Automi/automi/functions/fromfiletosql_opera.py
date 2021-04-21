import pandas as pd
import math
import sys
import os
import functions.mailsender as m


sys.path.insert(0, r"../.")

import var.sql as sql
import var.links_paths as lp
import update_mail as um

filename = "cleanOpeRAdf.csv"

# OpeRA converter from xlsx to 3 file for DB
def cleanerOP(table, rows, path, monthyear):
    cnx = sql.cnx
    crs = sql.createCursor()

# CLEANING EXCEL FILE

    df = pd.read_excel(path, skiprows=rows, dtype="object", skipfooter=1)
    x = 0
    check = 0
    for i in range(0, len(df.index)):
        if df["Unnamed: 0"][i] == "BACINO":
            check += 1
        if check > 1:
            df = df.drop([i])
        else:
            try:
                if math.isnan(df["Unnamed: 0"][i]):
                    x += 1
                    df["Unnamed: 0"][i] = df["Unnamed: 0"][i-x]
            except TypeError:
                x = 0

    if df["Unnamed: 1"][len(df.index)-1] == "TOTALE":
        None
    else:
        df = df.drop([len(df.index)-1])

# END CLEANING

# UPDATE GROUP AND SERVICE DB
    grouplist = list(df["Unnamed: 0"].unique())
    serviceslist = list(df["Unnamed: 1"].unique())

    group = pd.read_sql("SELECT * FROM ced_c87_group", cnx, index_col="ID")
    services = pd.read_sql("SELECT * FROM ced_OpeRA_services", cnx, index_col="ID")

    serviceslist = [i for i in serviceslist[1:] if i not in services.SERVICE.tolist()]
    grouplist = [i for i in grouplist[1:] if i not in group.BACINO.tolist()]

    change = []

    for i in range(0, len(services.index)):
        change.append(int(services.index[i]))
    services.index = change

    change = []
    for i in range(0, len(group.index)):
        change.append(int(group.index[i]))
    group.index = change

    for i in range(0,len(serviceslist)):
        query = ("INSERT INTO ced_opera_services (`ID`, `SERVICE`) VALUES ( %s, '%s')" % (str(i+max(services.index.tolist())+1), serviceslist[i]))
        crs.execute(query, cnx)
        cnx.commit()

    for i in range(0,len(grouplist)):
        query = ("INSERT INTO ced_c87_group (`ID`, `BACINO`) VALUES ( %s, '%s')" % (str(i+max(group.index.tolist())+1), grouplist[i]))
        crs.execute(query, cnx)
        cnx.commit()

    if len(grouplist) > 0:
        um.alertupd(grouplist, "OPERA")

    if len(serviceslist) > 0:
        um.alertupd(serviceslist, "OPERA")

    group = pd.read_sql("SELECT * FROM ced_c87_group", cnx, index_col="ID")
    services = pd.read_sql("SELECT * FROM ced_OpeRA_services", cnx, index_col="ID")
# END UPDATE GROUP AND SERVICE DB

# dfFRAME PREPARATION
    for i in range(0, len(df.index)):
        try:
            df.replace(df["Unnamed: 1"][i], services.index[services.SERVICE == df["Unnamed: 1"][i]].tolist()[0], inplace=True)
        except IndexError:
            None
        try:
            df.replace(df["Unnamed: 0"][i], group.index[group.BACINO == df["Unnamed: 0"][i]].tolist()[0], inplace=True)
        except IndexError:
            None

    service_aslist = list(df["Unnamed: 1"])
    group_aslist = list(df["Unnamed: 0"])
    for i in range(0, len(df.index)):
        service_aslist[i] = str(service_aslist[i]).zfill(3)
        group_aslist[i] = str(group_aslist[i]).zfill(3)

    df["Unnamed: 1"] = service_aslist
    df["Unnamed: 0"] = group_aslist

    columns_aslist = df.columns.tolist()
    new_index = "Keys"
    for i in range(0, len(columns_aslist)):
        if columns_aslist[i][:7] == "Unnamed":
            columns_aslist[i] = new_index + "_" + str(df[columns_aslist[i]][0])
        else:
            new_index = columns_aslist[i]
            columns_aslist[i] = new_index + "_" + str(df[columns_aslist[i]][0])
    df.columns = columns_aslist
    df = df.drop(0)

    df.insert(loc=0, column="MonthYear", value=monthyear)
    id = []
    for i in range(1, len(df.index)+1):
        id.append(df["MonthYear"][i] + df["Keys_BACINO"][i] + df["Keys_SERVIZIO"][i])

    df.insert(loc=0, column="ID", value=id)
    df.replace(to_replace=r"-", value="0", regex=True, inplace=True)

    try:
        df.drop(labels="Esiti chiusura On Field_nan", axis=1, inplace=True)
    except:
        None
# END dfFRAME PREPARATION
# UPDATE TABLE SINTESI
    df.to_csv(lp.uploadpath + "/" + filename, sep=";", index=False)

    query = sql.repwithcsv(lp.uploadpath + "/" + filename, table, "1")
    crs.execute(query)
    cnx.commit()

    crs.close()
    print(query)
    try:
        os.remove(lp.uploadpath + "/" + filename)
        access = open(lp.accesslogpath, "a")
        access.write("Dato caricato sulla tabella [ %s ].\n\r" % (table))
        access.close()
    except FileNotFoundError:
        errors = open(lp.errorlogpath, "a")
        errors.write("Impossibile trovare %s sulla cartella indicata - File non scaricato o illeggible \n\r"
                     % (lp.uploadpath + filename))
        errors.close()
        m.mail_error("OpeRA")
