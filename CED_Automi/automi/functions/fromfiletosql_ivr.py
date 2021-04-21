import xml.etree.ElementTree as ET
import pandas as pd
import sys
import os
from numpy import nan

sys.path.insert(0, r"../.")

import var.links_paths as lp
import var.sql as sql

crs = sql.createCursor()
cnx = sql.cnx


# from xml to excel IVR
def convert_xml(path, dbtable):
    dfcols = list()
    dflist = list()
    pathnew = lp.autodlpath + "\origin.xml"
    try:
        os.rename(path, pathnew)
        tree = ET.parse(pathnew)
    except FileNotFoundError:
        tree = ET.parse(pathnew)

    root = tree.getroot()
    i = 0
    for ws in root:
        for table in ws:
            for row in table:
                for cell in row:
                    for data in cell:
                        if i == 26:
                            dfcols.append(data.text)
                        if i > 26:
                            dflist.append(data.text)
                i += 1
                if i == 27:
                    df = pd.DataFrame(columns=dfcols)
                if i > 27:
                    df = df.append(pd.Series(dflist, index=dfcols), ignore_index=True)
                dflist = []

    df.fillna(value=0, inplace=True)
    df.to_csv(lp.uploadpath + "/convertedIVR.csv", sep=";", index=False)

    query = sql.repwithcsv(lp.uploadpath + "/convertedIVR.csv", dbtable, "1")
    crs.execute(query)
    cnx.commit()
    crs.close()
    try:
        os.remove(pathnew)
        try:
            os.remove(path)
        except FileNotFoundError:
            None
    except FileNotFoundError:
        None
    os.remove(lp.uploadpath + "/convertedIVR.csv")


filename = "\AssistenzaTecnica IVR.xls"
x = input("tabella ? ")
convert_xml(lp.autodlpath + filename, "ced_c87_ivr_%s" % x)
