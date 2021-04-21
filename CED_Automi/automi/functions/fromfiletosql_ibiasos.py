import pandas as pd
import sys
import os
import datetime as dt
import time


sys.path.insert(0, r"../.")

import var.sql as sql
import var.links_paths as lp
import functions.easter

crs = sql.createCursor()
cnx = sql.cnx

def convert_excel(anno, mese, giorno):
    isoday = list()
    days = ["LUNEDÌ", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ", "SABATO", "DOMENICA"]
    data = anno + "-" + mese + "-" + giorno
    dayofweek = days[dt.datetime.strptime(data, "%Y-%m-%d").isoweekday()-1]
    crs = sql.createCursor()
    query = sql.month_query(mese)
    crs.execute(query)
    x = 0
    mese = crs.fetchone()[0]
    query = "SELECT * FROM ced_festivity_tr WHERE monthname_id = '%s'" % mese
    crs.execute(query)
    fest = crs.fetchall()
    if mese == "03" or mese == "04":
        fest.append(easter.easter_monday(int(anno)))
    filename = "Dettagli_EN,LAV_IMPROPRIE," + giorno + mese.upper() + anno + dayofweek + ",ANNO" + anno + "," + mese.upper() + anno + ".xlsx"
    trying = True
    while trying:
        try:
            df = pd.read_excel(lp.autodlpath + "\\" + filename, skiprows=3, dtype="object")
            trying = False
        except FileNotFoundError:
            time.sleep(1)
            x += 1
            if x == 20:
                trying = False
    csvpath = lp.uploadpath + "/coverted.csv"
    df["DATA SEGNALAZIONE"] = df["DATA SEGNALAZIONE"].str[:16]
    df["DATA CHIUSURA"] = df["DATA CHIUSURA"].str[:16]
    df["DATA LAVORAZIONE"] = df["DATA LAVORAZIONE"].str[:16]
    df["DATA SOSPENSIONE ATPAY"] = df["DATA SOSPENSIONE ATPAY"].str[:16]
    df["DATA PRESA IN CARICO"] = df["DATA PRESA IN CARICO"].str[:16]
    df["DATA ASSEGNAZIONE ADDETTO"] = df["DATA ASSEGNAZIONE ADDETTO"].str[:16]
    query = sql.repwithcsv(csvpath, "ced_sos_easy_ibia", "1")
    group = pd.read_sql("SELECT * FROM ced_c87_group", cnx, index_col="ID")
    ascs = pd.read_sql("SELECT * FROM ced_c87_asc", cnx, index_col="ID")
    df = df.fillna(0)
    for i in range(0, len(df.index)):
        try:
            df["SEDE LAVORAZIONE"].replace(df["SEDE LAVORAZIONE"][i], ascs.index[ascs.ASC == df["SEDE LAVORAZIONE"][i]].tolist()[0], inplace=True)
        except IndexError:
            None
        try:
            df["TEAM LAVORAZIONE"].replace(df["TEAM LAVORAZIONE"][i], group.index[group.BACINO == df["TEAM LAVORAZIONE"][i]].tolist()[0], inplace=True)
        except IndexError:
            None
        if fest != []:
            for j in range(0, len(list(fest))):
                if str(df["DATA TK"][i])[6:] == str(fest[j][2]).zfill(2):
                    isoday.append(7)
                else:
                    isoday.append(dt.datetime.strptime(str(df["DATA TK"][i]), "%Y%m%d").isoweekday())
        else:
            isoday.append(dt.datetime.strptime(str(df["DATA TK"][i]), "%Y%m%d").isoweekday())

    df.insert(loc=len(df.columns), column="ISOWEEKDAY", value=isoday)
    ascs_aslist = list(df["SEDE LAVORAZIONE"])
    group_aslist = list(df["TEAM LAVORAZIONE"])
    for i in range(0, len(df.index)):
        ascs_aslist[i] = str(ascs_aslist[i]).zfill(2)
        group_aslist[i] = str(group_aslist[i]).zfill(3)

    df["SEDE LAVORAZIONE"] = ascs_aslist
    df["TEAM LAVORAZIONE"] = group_aslist

    df.to_csv(csvpath, sep=";", index=False)
    crs.execute(query)
    cnx.commit()
    os.remove(lp.autodlpath + "\\" + filename)
    os.remove(csvpath)

convert_excel("2018", "10", "31")
