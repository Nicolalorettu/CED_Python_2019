import pandas as pd
import datetime as dt
import numpy as np
from django.db import connection
import decimal
import warnings
import sys

sys.path.insert(0, r"../.")

try:
    import CED.var.query as qu
    import CED.var.targets as tz
except:
    import var.query as qu
    import var.targets as tz


# warnings.filterwarnings(category=RuntimeWarning, action="error")

def dfbasicop(dataframe, column, index1, index2, operation):
    if operation == "sum":
        result = dataframe[column][index1] + dataframe[column][index2]
    elif operation == "sub":
        result = dataframe[column][index1] - dataframe[column][index2]
    elif operation == "mol":
        result = dataframe[column][index1] * dataframe[column][index2]
    elif operation == "div":
        result = dataframe[column][index1] / dataframe[column][index2]
    return result


def dfkpiop(dataframe, column, type):
    # i type senza l'anno sono validi per tutti i contratti esistenti sino ad ora
    result = 0
    if type == "adsl":
        index = ["100", "103"]
    elif type == "adsl2018":
        index = ["101", "109", "104", "105", "110"]
    elif type == "fibra":
        index = ["102", "106"]
    elif type == "bofonia":
        index = ["103", "108", "111", "112"]
    elif type == "boadsl":
        index = ["101", "104", "105", "107", "109"]
    if type == "adsl":
        result = dataframe[column][index[0]] - dataframe[column][index[1]]
    else:
        for i in range(0, len(index)):
            try:
                result += dataframe[column][index[i]]
            except KeyError:
                None
    return result


def smscalcmonth(df, dfibia, month, columns):
    smstable = []
    mlist = list(month)
    for j in range(0, len(mlist)):
        monthnumber = month[j].number
        smstable.append(dict())
        smstable[j]["periodo"] = month[j].name
        # Fatturati Online
        smstable[j]["fatturati"] = dfibia["DATA_TK"][dfibia["DATA_TK"].str[4:6] == monthnumber].count()
        for i in range(0, 3):
            if i == 0:
                # Fake
                smstable[j]["fake"] = df[i][columns[i][1]][df[i]["Data_Invio"].str[5:7] == monthnumber].count()
                smstable[j]["percfake"] = round(decimal.Decimal(smstable[j]["fake"] / smstable[j]["fatturati"]*100), 2)
                if smstable[j]["percfake"] != smstable[j]["percfake"]:
                    smstable[j]["percfake"] = 0
            elif i == 1:
                # Non Presenti
                smstable[j]["Non_Presenti"] = df[i][columns[i][1]][df[i]["DATA_ETT"].str[5:7] == monthnumber].count()
            elif i == 2:
                # Totale, KO, totale univoco, ko univoco
                smstable[j]["Totale"] = df[i][columns[i][1]][df[i]["Data_Invio_Risposta_Cliente"].str[5:7] == monthnumber].count()
                smstable[j]["KO"] = df[i]["Esito"][(df[i]["Esito"] == "KO") & (df[i]["Data_Invio_Risposta_Cliente"].str[5:7] == monthnumber)].count()
                smstable[j]["percKO_Risp"] = round(decimal.Decimal(smstable[j]["KO"] / smstable[j]["Totale"]*100), 2)
                if smstable[j]["percKO_Risp"] != smstable[j]["percKO_Risp"]:
                    smstable[j]["percKO_Risp"] = 0
                dumpdf = df[i].drop_duplicates(subset=["Cod_Invio"])
                smstable[j]["Totale_TT"] = len(dumpdf["Cod_Invio"][(dumpdf["Data_Invio_Risposta_Cliente"].str[5:7] == monthnumber)])
                smstable[j]["KO_TT"] = len(dumpdf["Cod_Invio"][(dumpdf["Data_Invio_Risposta_Cliente"].str[5:7] == monthnumber) & (dumpdf["Esito"] == "KO")])
                if smstable[j]["Totale_TT"] != 0:
                    smstable[j]["percKO_TT"] = round(decimal.Decimal(smstable[j]["KO_TT"] / smstable[j]["Totale_TT"]*100), 2)
                else:
                    smstable[j]["percKO_TT"] = 0
    return smstable


def smscalcweek(df, dfibia, fweek, lweek, columns):
    smstable = []
    for j in range(fweek, lweek+1):
        smstable.append(dict())
        k = j-fweek
        smstable[k]["Week"] = str(j).zfill(2)
        # Fatturati Online
        smstable[k]["fatturati"] = dfibia["WEEK"][dfibia["WEEK"] == str(j)].count()
        for i in range(0, 3):
            if i == 0:
                # Fake
                smstable[k]["fake"] = df[i][columns[i][1]][df[i]["WEEK"] == str(j)].count()
                smstable[k]["percfake"] = round(decimal.Decimal(smstable[k]["fake"] / smstable[k]["fatturati"]*100), 2)
                if smstable[k]["percfake"] != smstable[k]["percfake"]:
                    smstable[k]["percfake"] = 0
            elif i == 1:
                # Non Presenti
                smstable[k]["Non_Presenti"] = df[i][columns[i][1]][df[i]["WEEK"] == str(j)].count()
            elif i == 2:
                # Totale, KO, totale univoco, ko univoco
                smstable[k]["Totale"] = df[i][columns[i][1]][df[i]["WEEK"] == str(j)].count()
                smstable[k]["KO"] = df[i]["Esito"][(df[i]["Esito"] == "KO") & (df[i]["WEEK"] == str(j))].count()
                smstable[k]["percKO_Risp"] = round(decimal.Decimal(smstable[k]["KO"] / smstable[k]["Totale"]*100), 2)
                if smstable[k]["percKO_Risp"] != smstable[k]["percKO_Risp"]:
                    smstable[k]["percKO_Risp"] = 0
                dumpdf = df[i].drop_duplicates(subset=["Cod_Invio"])
                smstable[k]["Totale_TT"] = len(dumpdf["Cod_Invio"][(dumpdf["WEEK"] == str(j))])
                smstable[k]["KO_TT"] = len(dumpdf["Cod_Invio"][(dumpdf["WEEK"] == str(j)) & (dumpdf["Esito"] == "KO")])
                if smstable[k]["Totale_TT"] != 0:
                    smstable[k]["percKO_TT"] = round(decimal.Decimal(smstable[k]["KO_TT"] / smstable[k]["Totale_TT"]*100), 2)
                else:
                    smstable[k]["percKO_TT"] = 0
    return smstable


def dfkpirich(dataframe, column, type):
    result = 0
    index = []
    if type == "semp872017":
        index = ["121", "120"]
    elif type == "semp872018":
        index = ["121", "120", "123", "124"]
    elif type == "mcmpl872017":
        index = ["122", "128", "124"]
    elif type == "mcmpl872018":
        index = ["122", "128", "126"]
    elif type == "cmpl872017":
        index = ["126", "125"]
    elif type == "cmpl1872018":
        index = ["125"]
    elif type == "1912017":
        index = ["130", "133", "131", "134", "132", "129"]
    elif type == "mcmpl1912018":
        index = ["130", "134", "132", "133", "129"]
    elif type == "cmpl1912018":
        index = ["131"]
    for i in index:
        try:
            result += dataframe[column][dataframe["Keys_SERVIZIO_id"] == i].sum()
        except KeyError:
            None
    return result


def KpiFEHome2017(df, columns):
    kpitable = list()
# kpitable["ReworkHF"]
    num = (df[columns[39]]["103"]+df[columns[41]]["103"])
    den = (df[columns[12]]["103"]+df[columns[33]]["103"])
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["ReworkHD"]
    num = dfkpiop(df, columns[39], "adsl")+dfkpiop(df, columns[41], "adsl")
    den = dfkpiop(df, columns[12], "adsl")+dfkpiop(df, columns[33], "adsl")
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioBOHF"]
    num = (df[columns[21]]["103"]+df[columns[23]]["103"]+df[columns[34]]["103"]+df[columns[35]]["103"]
           + df[columns[36]]["103"]+df[columns[37]]["103"])
    den = (df[columns[9]]["103"]+df[columns[12]]["103"]+df[columns[21]]["103"]+df[columns[22]]["103"]+df[columns[23]]["103"]
           + df[columns[24]]["103"]+df[columns[25]]["103"]+df[columns[35]]["103"]+df[columns[36]]["103"]+df[columns[37]]["103"])
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioBOHD"]
    num = (dfkpiop(df, columns[21], "adsl") + dfkpiop(df, columns[23], "adsl")
           + dfkpiop(df, columns[34], "adsl")+dfkpiop(df, columns[35], "adsl")
           + dfkpiop(df, columns[36], "adsl")+dfkpiop(df, columns[37], "adsl"))
    den = (dfkpiop(df, columns[9], "adsl")+dfkpiop(df, columns[12], "adsl")
           + dfkpiop(df, columns[21], "adsl")+dfkpiop(df, columns[22], "adsl")
           + dfkpiop(df, columns[23], "adsl")+dfkpiop(df, columns[24], "adsl")
           + dfkpiop(df, columns[25], "adsl")+dfkpiop(df, columns[35], "adsl")
           + dfkpiop(df, columns[36], "adsl")+dfkpiop(df, columns[37], "adsl"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioOFHF"]
    num = (df[columns[24]]["103"]+df[columns[25]]["103"])
    den = (df[columns[9]]["103"]+df[columns[12]]["103"]+df[columns[21]]["103"]+df[columns[22]]["103"]+df[columns[23]]["103"]
           + df[columns[24]]["103"]+df[columns[25]]["103"]+df[columns[35]]["103"]+df[columns[36]]["103"]+df[columns[37]]["103"])
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioOFHD"]
    num = dfkpiop(df, columns[24], "adsl") + dfkpiop(df, columns[25], "adsl")
    den = (dfkpiop(df, columns[9], "adsl")+dfkpiop(df, columns[12], "adsl")
           + dfkpiop(df, columns[21], "adsl")+dfkpiop(df, columns[22], "adsl")
           + dfkpiop(df, columns[23], "adsl")+dfkpiop(df, columns[24], "adsl")
           + dfkpiop(df, columns[25], "adsl")+dfkpiop(df, columns[35], "adsl")
           + dfkpiop(df, columns[36], "adsl")+dfkpiop(df, columns[37], "adsl"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
    kpitable.append(0)
    kpitable.append(0)
    for j in kpitable:
        if j != j:
            j = 0
    return kpitable


def KpiBOHoOf2017(df, df2, columns):
    kpitable = list()
# kpitable["ReworkBOF"]
    num1 = (dfkpiop(df, columns[5], "bofonia") + dfkpiop(df, columns[6], "bofonia"))
    den1 = dfkpiop(df, columns[4], "bofonia")
    num2 = (dfkpiop(df2, columns[5], "bofonia") + dfkpiop(df2, columns[6], "bofonia"))
    den2 = dfkpiop(df2, columns[4], "bofonia")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["ReworkBOA"]
    num1 = (dfkpiop(df, columns[5], "boadsl") + dfkpiop(df, columns[6], "boadsl"))
    den1 = dfkpiop(df, columns[4], "boadsl")
    num2 = (dfkpiop(df2, columns[5], "boadsl") + dfkpiop(df2, columns[6], "boadsl"))
    den2 = dfkpiop(df2, columns[4], "boadsl")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["ReworkBOFi"]
    num1 = (dfkpiop(df, columns[5], "fibra") + dfkpiop(df, columns[6], "fibra"))
    den1 = dfkpiop(df, columns[4], "fibra")
    num2 = (dfkpiop(df2, columns[5], "fibra") + dfkpiop(df2, columns[6], "fibra"))
    den2 = dfkpiop(df2, columns[4], "fibra")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["InvioBOOFHF"]
    num1 = (dfkpiop(df, columns[13], "bofonia") + dfkpiop(df, columns[21], "bofonia"))
    den1 = dfkpiop(df, columns[3], "bofonia")
    num2 = (dfkpiop(df2, columns[13], "bofonia") + dfkpiop(df2, columns[21], "bofonia"))
    den2 = dfkpiop(df2, columns[3], "bofonia")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["InvioBOOFHA"]
    num1 = (dfkpiop(df, columns[13], "boadsl") + dfkpiop(df, columns[21], "boadsl"))
    den1 = dfkpiop(df, columns[3], "boadsl")
    num2 = (dfkpiop(df2, columns[13], "boadsl") + dfkpiop(df2, columns[21], "boadsl"))
    den2 = dfkpiop(df2, columns[3], "boadsl")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["InvioBOOFHFi"]
    num1 = (dfkpiop(df, columns[13], "fibra") + dfkpiop(df, columns[21], "fibra"))
    den1 = dfkpiop(df, columns[3], "fibra")
    num2 = (dfkpiop(df2, columns[13], "fibra") + dfkpiop(df2, columns[21], "fibra"))
    den2 = dfkpiop(df2, columns[3], "fibra")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["BOCTFD"]
    num1 = df[columns[11]]["100"]
    den1 = df[columns[3]]["100"]
    num2 = df2[columns[11]]["100"]
    den2 = df2[columns[3]]["100"]
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
    for j in kpitable:
        if j != j:
            j = 0
    return kpitable


def KpiFEOffice2017(df, columns):
    kpitable = list()
# kpitable["ReworkOF"]
    num = (df[columns[39]]["100"]+df[columns[41]]["100"])
    den = (df[columns[12]]["100"]+df[columns[33]]["100"])
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioBOOF"]
    num = (df[columns[21]]["100"]+df[columns[23]]["100"]+df[columns[34]]["100"]+df[columns[35]]["100"]
           + df[columns[36]]["100"]+df[columns[37]]["100"])
    den = (df[columns[9]]["100"]+df[columns[12]]["100"]+df[columns[21]]["100"]+df[columns[22]]["100"]+df[columns[23]]["100"]
           + df[columns[24]]["100"]+df[columns[25]]["100"]+df[columns[35]]["100"]+df[columns[36]]["100"]+df[columns[37]]["100"])
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioOFOF"]
    num = (df[columns[24]]["100"]+df[columns[25]]["100"])
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
    for j in kpitable:
        if j != j:
            j = 0
    return kpitable


def KpiFEHome2018(df, columns):
    kpitable = list()
# kpitable["ReworkHF"]
    num = (df[columns[39]]["103"]+df[columns[41]]["103"])
    den = (df[columns[12]]["103"]+df[columns[35]]["103"])
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["ReworkHD"]
    num = dfkpiop(df, columns[39], "adsl2018")+dfkpiop(df, columns[41], "adsl2018")
    den = dfkpiop(df, columns[12], "adsl2018")+dfkpiop(df, columns[35], "adsl2018")
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["ReworkHFi"]
    num = dfkpiop(df, columns[39], "fibra")+dfkpiop(df, columns[41], "fibra")
    den = dfkpiop(df, columns[12], "fibra")+dfkpiop(df, columns[35], "fibra")
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["ONTHF"]
    num = df[columns[75]]["103"]
    den = df[columns[24]]["103"] + df[columns[25]]["103"]
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["ONTHD"]
    num = dfkpiop(df, columns[75], "adsl2018")
    den = dfkpiop(df, columns[24], "adsl2018")+dfkpiop(df, columns[25], "adsl2018")
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["ONTHFi"]
    num = dfkpiop(df, columns[75], "fibra")
    den = dfkpiop(df, columns[24], "fibra")+dfkpiop(df, columns[25], "fibra")
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioBOHF"]
    num = (df[columns[21]]["103"]+df[columns[23]]["103"]+df[columns[34]]["103"]+df[columns[35]]["103"]
           + df[columns[36]]["103"]+df[columns[37]]["103"])
    den = (df[columns[9]]["103"]+df[columns[12]]["103"]+df[columns[21]]["103"]+df[columns[22]]["103"]+df[columns[23]]["103"]
           + df[columns[24]]["103"]+df[columns[25]]["103"]+df[columns[35]]["103"]+df[columns[36]]["103"]+df[columns[37]]["103"])
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))

# kpitable["InvioBOHD"]
    num = (dfkpiop(df, columns[21], "adsl2018") + dfkpiop(df, columns[23], "adsl2018")
           + dfkpiop(df, columns[34], "adsl2018")+dfkpiop(df, columns[35], "adsl2018")
           + dfkpiop(df, columns[36], "adsl2018")+dfkpiop(df, columns[37], "adsl2018"))
    den = (dfkpiop(df, columns[9], "adsl2018")+dfkpiop(df, columns[12], "adsl2018")
           + dfkpiop(df, columns[21], "adsl2018")+dfkpiop(df, columns[22], "adsl2018")
           + dfkpiop(df, columns[23], "adsl2018")+dfkpiop(df, columns[24], "adsl2018")
           + dfkpiop(df, columns[25], "adsl2018")+dfkpiop(df, columns[35], "adsl2018")
           + dfkpiop(df, columns[36], "adsl2018")+dfkpiop(df, columns[37], "adsl2018"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioBOHFi"]
    num = (dfkpiop(df, columns[21], "fibra") + dfkpiop(df, columns[23], "fibra")
           + dfkpiop(df, columns[34], "fibra")+dfkpiop(df, columns[35], "fibra")
           + dfkpiop(df, columns[36], "fibra")+dfkpiop(df, columns[37], "fibra"))
    den = (dfkpiop(df, columns[9], "fibra")+dfkpiop(df, columns[12], "fibra")
           + dfkpiop(df, columns[21], "fibra")+dfkpiop(df, columns[22], "fibra")
           + dfkpiop(df, columns[23], "fibra")+dfkpiop(df, columns[24], "fibra")
           + dfkpiop(df, columns[25], "fibra")+dfkpiop(df, columns[35], "fibra")
           + dfkpiop(df, columns[36], "fibra")+dfkpiop(df, columns[37], "fibra"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioOFHF"]
    num = (df[columns[24]]["103"]+df[columns[25]]["103"])
    den = (df[columns[9]]["103"]+df[columns[12]]["103"]+df[columns[21]]["103"]+df[columns[22]]["103"]+df[columns[23]]["103"]
           + df[columns[24]]["103"]+df[columns[25]]["103"]+df[columns[35]]["103"]+df[columns[36]]["103"]+df[columns[37]]["103"])
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioOFHD"]
    num = dfkpiop(df, columns[24], "adsl2018") + dfkpiop(df, columns[25], "adsl2018")
    den = (dfkpiop(df, columns[9], "adsl2018")+dfkpiop(df, columns[12], "adsl2018")
           + dfkpiop(df, columns[21], "adsl2018")+dfkpiop(df, columns[22], "adsl2018")
           + dfkpiop(df, columns[23], "adsl2018")+dfkpiop(df, columns[24], "adsl2018")
           + dfkpiop(df, columns[25], "adsl2018")+dfkpiop(df, columns[35], "adsl2018")
           + dfkpiop(df, columns[36], "adsl2018")+dfkpiop(df, columns[37], "adsl2018"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioOFHFi"]
    num = dfkpiop(df, columns[24], "fibra") + dfkpiop(df, columns[25], "fibra")
    den = (dfkpiop(df, columns[9], "fibra")+dfkpiop(df, columns[12], "fibra")
           + dfkpiop(df, columns[21], "fibra")+dfkpiop(df, columns[22], "fibra")
           + dfkpiop(df, columns[23], "fibra")+dfkpiop(df, columns[24], "fibra")
           + dfkpiop(df, columns[25], "fibra")+dfkpiop(df, columns[35], "fibra")
           + dfkpiop(df, columns[36], "fibra")+dfkpiop(df, columns[37], "fibra"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
    for j in kpitable:
        if j != j:
            j = 0
    return kpitable


def KpiFEOffice2018(df, columns):
    kpitable = list()
# kpitable["ReworkOFM"]
    num = (dfkpiop(df, columns[39], "bofonia") + dfkpiop(df, columns[41], "bofonia") +
           dfkpiop(df, columns[39], "adsl2018") + dfkpiop(df, columns[41], "adsl2018"))
    den = (dfkpiop(df, columns[12], "bofonia")+dfkpiop(df, columns[35], "bofonia") +
           dfkpiop(df, columns[12], "adsl2018")+dfkpiop(df, columns[35], "adsl2018"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["ReworkOFC"]
    num = dfkpiop(df, columns[39], "fibra") + dfkpiop(df, columns[41], "fibra")
    den = dfkpiop(df, columns[12], "fibra")+dfkpiop(df, columns[35], "fibra")
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["ONTMOF"]
    num = dfkpiop(df, columns[75], "bofonia") + dfkpiop(df, columns[75], "adsl2018")
    den = (dfkpiop(df, columns[24], "bofonia")+dfkpiop(df, columns[25], "bofonia") +
           dfkpiop(df, columns[24], "adsl2018")+dfkpiop(df, columns[25], "adsl2018"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["ONTCOF"]
    num = dfkpiop(df, columns[75], "fibra")
    den = dfkpiop(df, columns[24], "fibra")+dfkpiop(df, columns[25], "fibra")
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioBOMOF"]
    num = ((dfkpiop(df, columns[21], "bofonia") + dfkpiop(df, columns[23], "bofonia") + dfkpiop(df, columns[34], "bofonia") +
           dfkpiop(df, columns[35], "bofonia") + dfkpiop(df, columns[36], "bofonia") + dfkpiop(df, columns[37], "bofonia")) +
           (dfkpiop(df, columns[21], "adsl2018") + dfkpiop(df, columns[23], "adsl2018") + dfkpiop(df, columns[34], "adsl2018") +
           dfkpiop(df, columns[35], "adsl2018") + dfkpiop(df, columns[36], "adsl2018") + dfkpiop(df, columns[37], "adsl2018")))
    den = ((dfkpiop(df, columns[9], "bofonia") + dfkpiop(df, columns[12], "bofonia") + dfkpiop(df, columns[21], "bofonia") +
           dfkpiop(df, columns[22], "bofonia") + dfkpiop(df, columns[23], "bofonia") + dfkpiop(df, columns[24], "bofonia") +
           dfkpiop(df, columns[25], "bofonia") + dfkpiop(df, columns[35], "bofonia") + dfkpiop(df, columns[36], "bofonia") +
           dfkpiop(df, columns[37], "bofonia")) +
           (dfkpiop(df, columns[9], "adsl2018") + dfkpiop(df, columns[12], "adsl2018") + dfkpiop(df, columns[21], "adsl2018") +
            dfkpiop(df, columns[22], "adsl2018") + dfkpiop(df, columns[23], "adsl2018") + dfkpiop(df, columns[24], "adsl2018") +
            dfkpiop(df, columns[25], "adsl2018") + dfkpiop(df, columns[35], "adsl2018") + dfkpiop(df, columns[36], "adsl2018") +
            dfkpiop(df, columns[37], "adsl2018")))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioBOCOF"]
    num = (dfkpiop(df, columns[21], "fibra") + dfkpiop(df, columns[23], "fibra") + dfkpiop(df, columns[34], "fibra") +
           dfkpiop(df, columns[35], "fibra") + dfkpiop(df, columns[36], "fibra") + dfkpiop(df, columns[37], "fibra"))
    den = (dfkpiop(df, columns[9], "fibra") + dfkpiop(df, columns[12], "fibra") + dfkpiop(df, columns[21], "fibra") +
           dfkpiop(df, columns[22], "fibra") + dfkpiop(df, columns[23], "fibra") + dfkpiop(df, columns[24], "fibra") +
           dfkpiop(df, columns[25], "fibra") + dfkpiop(df, columns[35], "fibra") + dfkpiop(df, columns[36], "fibra") +
           dfkpiop(df, columns[37], "fibra"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioOFMOF"]
    num = (dfkpiop(df, columns[24], "bofonia") + dfkpiop(df, columns[25], "bofonia") +
           dfkpiop(df, columns[24], "adsl2018") + dfkpiop(df, columns[25], "adsl2018"))
    den = ((dfkpiop(df, columns[9], "bofonia") + dfkpiop(df, columns[12], "bofonia") + dfkpiop(df, columns[21], "bofonia") +
           dfkpiop(df, columns[22], "bofonia") + dfkpiop(df, columns[23], "bofonia") + dfkpiop(df, columns[24], "bofonia") +
           dfkpiop(df, columns[25], "bofonia") + dfkpiop(df, columns[35], "bofonia") + dfkpiop(df, columns[36], "bofonia") +
           dfkpiop(df, columns[37], "bofonia")) +
           (dfkpiop(df, columns[9], "adsl2018") + dfkpiop(df, columns[12], "adsl2018") + dfkpiop(df, columns[21], "adsl2018") +
            dfkpiop(df, columns[22], "adsl2018") + dfkpiop(df, columns[23], "adsl2018") + dfkpiop(df, columns[24], "adsl2018") +
            dfkpiop(df, columns[25], "adsl2018") + dfkpiop(df, columns[35], "adsl2018") + dfkpiop(df, columns[36], "adsl2018") +
            dfkpiop(df, columns[37], "adsl2018")))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
# kpitable["InvioOFCOF"]
    num = dfkpiop(df, columns[24], "fibra") + dfkpiop(df, columns[25], "fibra")
    den = (dfkpiop(df, columns[9], "fibra") + dfkpiop(df, columns[12], "fibra") + dfkpiop(df, columns[21], "fibra") +
           dfkpiop(df, columns[22], "fibra") + dfkpiop(df, columns[23], "fibra") + dfkpiop(df, columns[24], "fibra") +
           dfkpiop(df, columns[25], "fibra") + dfkpiop(df, columns[35], "fibra") + dfkpiop(df, columns[36], "fibra") +
           dfkpiop(df, columns[37], "fibra"))
    kpitable.append(round(decimal.Decimal(num / den * 100), 2))
    for j in kpitable:
        if j != j:
            j = 0
    return kpitable


def KpiBOHoOf2018(df, df2, columns):
    kpitable = list()
# kpitable["ReworkBOF"]
    num1 = (dfkpiop(df, columns[5], "bofonia") + dfkpiop(df, columns[6], "bofonia"))
    den1 = dfkpiop(df, columns[4], "bofonia")
    num2 = (dfkpiop(df2, columns[5], "bofonia") + dfkpiop(df2, columns[6], "bofonia"))
    den2 = dfkpiop(df2, columns[4], "bofonia")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["ReworkBOD"]
    num1 = (dfkpiop(df, columns[5], "boadsl") + dfkpiop(df, columns[6], "boadsl"))
    den1 = dfkpiop(df, columns[4], "boadsl")
    num2 = (dfkpiop(df2, columns[5], "boadsl") + dfkpiop(df2, columns[6], "boadsl"))
    den2 = dfkpiop(df2, columns[4], "boadsl")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["ReworkBOFi"]
    num1 = (dfkpiop(df, columns[5], "fibra") + dfkpiop(df, columns[6], "fibra"))
    den1 = dfkpiop(df, columns[4], "fibra")
    num2 = (dfkpiop(df2, columns[5], "fibra") + dfkpiop(df2, columns[6], "fibra"))
    den2 = dfkpiop(df2, columns[4], "fibra")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["ONTBOF"]
    num1 = dfkpiop(df, columns[27], "bofonia")
    den1 = dfkpiop(df, columns[13], "bofonia") + dfkpiop(df, columns[21], "bofonia")
    num2 = dfkpiop(df2, columns[27], "bofonia")
    den2 = dfkpiop(df2, columns[13], "bofonia") + dfkpiop(df, columns[21], "bofonia")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["ONTBOD"]
    num1 = dfkpiop(df, columns[27], "boadsl")
    den1 = dfkpiop(df, columns[13], "boadsl") + dfkpiop(df, columns[21], "boadsl")
    num2 = dfkpiop(df2, columns[27], "boadsl")
    den2 = dfkpiop(df2, columns[13], "boadsl") + dfkpiop(df, columns[21], "boadsl")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["ONTBOFi"]
    num1 = dfkpiop(df, columns[27], "fibra")
    den1 = dfkpiop(df, columns[13], "fibra") + dfkpiop(df, columns[21], "fibra")
    num2 = dfkpiop(df2, columns[27], "fibra")
    den2 = dfkpiop(df2, columns[13], "fibra") + dfkpiop(df, columns[21], "fibra")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["InvioBOOFHF"]
    num1 = (dfkpiop(df, columns[13], "bofonia") + dfkpiop(df, columns[21], "bofonia"))
    den1 = dfkpiop(df, columns[3], "bofonia")
    num2 = (dfkpiop(df2, columns[13], "bofonia") + dfkpiop(df2, columns[21], "bofonia"))
    den2 = dfkpiop(df2, columns[3], "bofonia")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["InvioBOOFHD"]
    num1 = (dfkpiop(df, columns[13], "boadsl") + dfkpiop(df, columns[21], "boadsl"))
    den1 = dfkpiop(df, columns[3], "boadsl")
    num2 = (dfkpiop(df2, columns[13], "boadsl") + dfkpiop(df2, columns[21], "boadsl"))
    den2 = dfkpiop(df2, columns[3], "boadsl")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["InvioBOOFHFi"]
    num1 = (dfkpiop(df, columns[13], "fibra") + dfkpiop(df, columns[21], "fibra"))
    den1 = dfkpiop(df, columns[3], "fibra")
    num2 = (dfkpiop(df2, columns[13], "fibra") + dfkpiop(df2, columns[21], "fibra"))
    den2 = dfkpiop(df2, columns[3], "fibra")
    kpitable.append(round(decimal.Decimal((num1 + num2) / (den1 + den2) * 100), 2))
# kpitable["CTEAM"] + ["RIPETIZIONE 3gg"]
    kpitable.append(0)
    kpitable.append(0)
    for j in kpitable:
        if j != j:
            j = 0
    return kpitable


def ivrmonth(df, table, date, group):
    ivr = dict()
    ivrlist = list()
    segmento = "FIBRA"
    index = ["Periodo", "Gruppo", "Interviste_Valide", "Interviste_NV", "Media_DD4", "Media_DD7"]
    indexvoto = ["Voto_1", "Voto_2", "Voto_3", "Voto_4", "Voto_5", "Voto_6", "Voto_7", "Voto_8", "Voto_9", "Voto_10"]
    indextable = []
    if table == "ced_c87_ivr_187":
        indextable = ["Media_Fonia", "Media_Adsl"]
        macroesigenza = "FONIA"
        indexfibra = "Media_Fibra"
    elif table == "ced_c87_ivr_191":
        indextable = ["Media_NGAN", "Media_Fonia_Adsl"]
        macroesigenza = "NGAN"
    for i in range(0, len(indextable)):
        index.append(indextable[i])
    for i in range(0, len(indexvoto)):
        index.append(indexvoto[i])
    if table == "ced_c87_ivr_187":
        index.append(indexfibra)
    for month in date:
        ivr[index[0]] = month
        ivr[index[1]] = group
        ivr[index[2]] = len(df[(df["DD7"] > 0) & (df.DATA_INT.str[3:] == month)])
        ivr[index[3]] = len(df[(df["DD7"] == 0) & (df.DATA_INT.str[3:] == month)])
        ivr[index[4]] = round(decimal.Decimal(df[(df["DD7"] > 0) & (df.DATA_INT.str[3:] == month)]["DD4"].sum() / ivr["Interviste_Valide"]), 2)
        ivr[index[5]] = round(decimal.Decimal(df[df.DATA_INT.str[3:] == month]["DD7"].sum() / ivr["Interviste_Valide"]), 2)
        ivr[index[6]] = round(decimal.Decimal(df[(df["MACRO_ESIGENZA"] == macroesigenza) & (df.DATA_INT.str[3:] == month)]["DD7"].sum()
                             / len(df[(df["DD7"] > 0) & (df["MACRO_ESIGENZA"] == macroesigenza) & (df.DATA_INT.str[3:] == month)])), 2)
        ivr[index[7]] = round(decimal.Decimal(df[(df["MACRO_ESIGENZA"] != macroesigenza) & (df["SEGMENTO"] != segmento) & (df.DATA_INT.str[3:] == month)]["DD7"].sum()
                            / len(df[(df["DD7"] > 0) & (df["MACRO_ESIGENZA"] != macroesigenza) & (df["SEGMENTO"] != segmento) & (df.DATA_INT.str[3:] == month)])), 2)
        for j in range(8, 18):
            ivr[index[j]] = round(decimal.Decimal(len(df[(df["DD7"] == j-7) & (df.DATA_INT.str[3:] == month)])), 2)
        if table == "ced_c87_ivr_187":
            ivr[indexfibra] = round(decimal.Decimal(df[(df["SEGMENTO"] == segmento) & (df.DATA_INT.str[3:] == month)]["DD7"].sum()
                                      / len(df[(df["DD7"] > 0) & (df["SEGMENTO"] == segmento) & (df.DATA_INT.str[3:] == month)])), 2)
        for j in ivr:
            if ivr[j] != ivr[j]:
                ivr[j] = 0
        ivrlist.append(ivr.copy())
    return ivrlist


def richmonth(df, table, data, group, columns, contract):
    rich = dict()
    richlist = list()
    indextable = []
    mdict = dict()
    index = ["Periodo", "Gruppo", "Tipo", "Chiamate", "Abbattute", "Rich_3_gg", "Rich_3_gg_p"]
    if table == "ced_paco_richiamate_fe_187":
        if contract[0].id == 1:
            indextable = ["semp872017", "mcmpl872017", "cmpl872017"]
        elif contract[0].id == 4:
            indextable = ["semp872018", "mcmpl872018", "cmpl1872018"]
        indexhead = ["Semplici", "M.Compl.", "Complesse"]
    elif table == "ced_paco_richiamate_fe_191":
        if contract[0].id == 1:
            indextable = ["1912017"]
            indexhead = ["Office"]
        elif contract[0].id == 4:
            indextable = ["mcmpl1912018", "cmpl1912018"]
            indexhead = ["M.Compl.", "Complesse"]
    for mese in data:
        for j in range(0, len(indextable)): # Conto Mese Tipo
            rich[index[0]] = mese
            rich[index[1]] = group
            rich[index[2]] = indexhead[j]
            rich[index[3]] = dfkpirich(df[df.Data.str[3:8] == mese], columns[4], indextable[j])
            rich[index[4]] = dfkpirich(df[df.Data.str[3:8] == mese], columns[5], indextable[j])
            rich[index[5]] = dfkpirich(df[df.Data.str[3:8] == mese], columns[8], indextable[j])
            rich[index[6]] = round(decimal.Decimal(((rich[index[5]] / rich[index[3]])*100)), 2)
            for z in rich:
                if rich[z] != rich[z]:
                    rich[z] = 0
            richlist.append(rich.copy())
    return richlist


def richweek(df, table, fweek, lweek, years, group, columns, contract):
    rich = dict()
    richlist = list()
    indextable = []
    index = ["Periodo", "Gruppo", "Tipo", "Chiamate", "Abbattute", "Rich_3_gg", "Rich_3_gg_p"]
    try:
        df.insert(loc=1, column="WEEK", value=df["Data"])
        df["WEEK"] = pd.to_datetime(df["WEEK"], dayfirst=True)
        df["WEEK"] = df["WEEK"].dt.strftime("%W")
    except ValueError:
        None
    if table == "ced_paco_richiamate_fe_187":
        if contract[0].id == 1:
            indextable = ["semp872017", "mcmpl872017", "cmpl872017"]
        elif contract[0].id == 4:
            indextable = ["semp872018", "mcmpl872018", "cmpl1872018"]
        indexhead = ["Semplici", "M.Compl.", "Complesse"]
    elif table == "ced_paco_richiamate_fe_191":
        if contract[0].id == 1:
            indextable = ["1912017"]
            indexhead = ["Office"]
        elif contract[0].id == 4:
            indextable = ["mcmpl1912018", "cmpl1912018"]
            indexhead = ["M.Compl.", "Complesse"]
    endweek = 0
    for year in years:
        if lweek < fweek:
            endweek = lweek
            lweek = 52
        elif endweek > 0:
            fweek = 1
            lweek = endweek
        for week in range(fweek, lweek+1):
            for i in range(0, len(indextable)):
                rich[index[0]] = week
                rich[index[1]] = group
                rich[index[2]] = indexhead[i]
                rich[index[3]] = dfkpirich(df[(df.WEEK == str(week)) & (df.Data.str[6:] == str(year[2:]))], columns[4], indextable[i])
                rich[index[4]] = dfkpirich(df[(df.WEEK == str(week)) & (df.Data.str[6:] == str(year[2:]))], columns[5], indextable[i])
                rich[index[5]] = dfkpirich(df[(df.WEEK == str(week)) & (df.Data.str[6:] == str(year[2:]))], columns[8], indextable[i])
                rich[index[6]] = round(decimal.Decimal(((rich[index[5]] / rich[index[3]])*100)), 2)
                for z in rich:
                    if rich[z] != rich[z]:
                        rich[z] = 0
                richlist.append(rich.copy())
    return richlist


def ivrweek(df, table, fweek, lweek, years, group):
    ivr = dict()
    ivrlist = list()
    segmento = "FIBRA"
    index = ["Periodo", "Gruppo", "Interviste_Valide", "Interviste_NV", "Media_DD4", "Media_DD7"]
    indexvoto = ["Voto_1", "Voto_2", "Voto_3", "Voto_4", "Voto_5", "Voto_6", "Voto_7", "Voto_8", "Voto_9", "Voto_10"]
    indextable = []
    try:
        df.insert(loc=1, column="WEEK", value=df["DATA_INT"])
        df["WEEK"] = pd.to_datetime(df["WEEK"], dayfirst=True)
        df["WEEK"] = df["WEEK"].dt.strftime("%W")
    except ValueError:
        None
    if table == "ced_c87_ivr_187":
        indextable = ["Media_Fonia", "Media_Adsl"]
        macroesigenza = "FONIA"
        indexfibra = "Media_Fibra"
    elif table == "ced_c87_ivr_191":
        indextable = ["Media_NGAN", "Media_Fonia_Adsl"]
        macroesigenza = "NGAN"
    for i in range(0, len(indextable)):
        index.append(indextable[i])
    for i in range(0, len(indexvoto)):
        index.append(indexvoto[i])
    if table == "ced_c87_ivr_187":
        index.append(indexfibra)
    endweek = 0
    for year in years:
        if lweek < fweek:
            endweek = lweek
            lweek = 52
        elif endweek > 0:
            fweek = 1
            lweek = endweek
        for week in range(fweek, lweek+1):
            ivr[index[0]] = week
            ivr[index[1]] = group
            ivr[index[2]] = len(df[(df["DD7"] > 0) & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)])
            ivr[index[3]] = len(df[(df["DD7"] == 0) & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)])
            ivr[index[4]] = round(df[(df["DD7"] > 0) & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)]["DD4"].sum() / ivr["Interviste_Valide"], 2)
            ivr[index[5]] = round(df[df.WEEK == str(week)]["DD7"].sum() / ivr["Interviste_Valide"], 2)
            ivr[index[6]] = round((df[(df["MACRO_ESIGENZA"] == macroesigenza) & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)]["DD7"].sum()
                            / len(df[(df["DD7"] > 0) & (df["MACRO_ESIGENZA"] == macroesigenza) & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)])), 2)
            ivr[index[7]] = round((df[(df["MACRO_ESIGENZA"] != macroesigenza) & (df["SEGMENTO"] != segmento) & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)]["DD7"].sum()
                            / len(df[(df["DD7"] > 0) & (df["MACRO_ESIGENZA"] != macroesigenza) & (df["SEGMENTO"] != segmento) & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)])), 2)
            for j in range(8, 18):
                ivr[index[j]] = len(df[(df["DD7"] == j-7) & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)])
            if table == "ced_c87_ivr_187":
                ivr[indexfibra] = round((df[(df["SEGMENTO"] == segmento)  & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)]["DD7"].sum()
                                     / len(df[(df["DD7"] > 0) & (df["SEGMENTO"] == segmento) & (df.WEEK == str(week)) & (df.DATA_INT.str[6:] == year)])), 2)
            for j in ivr:
                if ivr[j] != ivr[j]:
                    ivr[j] = 0
            ivrlist.append(ivr.copy())
    return ivrlist


def ibia(df, group, today, rowlist, target, table):
    ibiadict = list()
    for i in range(0, len(list(rowlist))):
        ibiadict.append(dict())
        ibiadict[i]["date"] = (today)
        ibiadict[i]["group"] = (group[0].BACINO)
        ibiadict[i]["rowdesc"] = (rowlist[i].ROW)
        if i == 5:
            ibiadict[i]["FONIA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == rowlist[i+1].ID)  & (df["Keys_SERVIZIO_id"] == "13")]["Nro_TK"].sum() /
                                     df[(df["Keys_SERVIZIO_id"] == "13")]["Nro_TK"].sum()*100))
            ibiadict[i]["FONIATZ"] = target[table]["FONIA"][i]
            ibiadict[i]["ADSL"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == rowlist[i+1].ID)  & (df["Keys_SERVIZIO_id"] != "13")]["Nro_TK"].sum() /
                                    df[(df["Keys_SERVIZIO_id"] != "13")]["Nro_TK"].sum()*100))
            ibiadict[i]["ADSLTZ"] = target[table]["ADSL"][i]
            ibiadict[i]["FIBRA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == rowlist[i+1].ID)  & (df["Keys_SERVIZIO_id"] != "13")]["Nro_TK"].sum() /
                                         df[(df["Keys_SERVIZIO_id"] != "13")]["Nro_TK"].sum()*100))
            ibiadict[i]["FIBRATZ"] = target[table]["FIBRA"][i]
        elif i == 7:
            ibiadict[i]["FONIA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == rowlist[i-1].ID)  & (df["Keys_SERVIZIO_id"] == "13")]["Nro_TK"].sum() /
                                     df[(df["Keys_SERVIZIO_id"] == "13")]["Nro_TK"].sum()*100))
            ibiadict[i]["FONIATZ"] = target[table]["FONIA"][i]
            ibiadict[i]["ADSL"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == rowlist[i-1].ID)  & (df["Keys_SERVIZIO_id"] != "13")]["Nro_TK"].sum() /
                                    df[(df["Keys_SERVIZIO_id"] != "13")]["Nro_TK"].sum()*100))
            ibiadict[i]["ADSLTZ"] = target[table]["ADSL"][i]
            ibiadict[i]["FIBRA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == rowlist[i-1].ID)  & (df["Keys_SERVIZIO_id"] != "13")]["Nro_TK"].sum() /
                                         df[(df["Keys_SERVIZIO_id"] != "13")]["Nro_TK"].sum()*100))
            ibiadict[i]["FIBRATZ"] = target[table]["FIBRA"][i]
        else:
            ibiadict[i]["FONIA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == rowlist[i].ID)  & (df["Keys_SERVIZIO_id"] == "13")]["Nro_TK"].sum() /
                                     df[(df["Keys_SERVIZIO_id"] == "13")]["Nro_TK"].sum()*100))
            ibiadict[i]["FONIATZ"] = target[table]["FONIA"][i]
            ibiadict[i]["ADSL"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == rowlist[i].ID)  & (df["Keys_SERVIZIO_id"] == "10")]["Nro_TK"].sum() /
                                    df[(df["Keys_SERVIZIO_id"] == "10")]["Nro_TK"].sum()*100))
            ibiadict[i]["ADSLTZ"] = target[table]["ADSL"][i]
            ibiadict[i]["FIBRA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == rowlist[i].ID)  & (df["Keys_SERVIZIO_id"] == "11")]["Nro_TK"].sum() /
                                         df[(df["Keys_SERVIZIO_id"] == "11")]["Nro_TK"].sum()*100))
            ibiadict[i]["FIBRATZ"] = target[table]["FIBRA"][i]
    ibiadict.append(dict())
    ibiadict[i+1]["date"] = (today)
    ibiadict[i+1]["group"] = (group[0].BACINO)
    ibiadict[i+1]["rowdesc"] = ("VERIFICHE NEGATIVE")
    ibiadict[i+1]["FONIA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "13")]["Verifiche_Negative"].sum() /
                             df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "13")]["Nro_TK"].sum()*100))
    ibiadict[i+1]["FONIATZ"] = target[table]["FONIA"][i+1]
    ibiadict[i+1]["ADSL"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "10")]["Verifiche_Negative"].sum() /
                             df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "10")]["Nro_TK"].sum()*100))
    ibiadict[i+1]["ADSLTZ"] = target[table]["ADSL"][i+1]
    ibiadict[i+1]["FIBRA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "11")]["Verifiche_Negative"].sum() /
                             df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "11")]["Nro_TK"].sum()*100))
    ibiadict[i+1]["FIBRATZ"] = target[table]["FIBRA"][i+1]
    ibiadict.append(dict())
    ibiadict[i+2]["date"] = (today)
    ibiadict[i+2]["group"] = (group[0].BACINO)
    ibiadict[i+2]["rowdesc"] = ("VERIFICHE NEGATIVE VOL")
    ibiadict[i+2]["FONIA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "13")]["Verifiche_Negative"].sum()))
    ibiadict[i+2]["FONIATZ"] = target[table]["FONIA"][i+2]
    ibiadict[i+2]["ADSL"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "10")]["Verifiche_Negative"].sum()))
    ibiadict[i+2]["ADSLTZ"] = target[table]["ADSL"][i+2]
    ibiadict[i+2]["FIBRA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "11")]["Verifiche_Negative"].sum()))
    ibiadict[i+2]["FIBRATZ"] = target[table]["FIBRA"][i+2]
    ibiadict.append(dict())
    ibiadict[i+3]["date"] = (today)
    ibiadict[i+3]["group"] = (group[0].BACINO)
    ibiadict[i+3]["rowdesc"] = ("RIP 1GG")
    ibiadict[i+3]["FONIA"] = (((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "13")]["Nro_TK_RIP_1GG"].sum() +
                                df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "13")]["Verifiche_Negative"].sum())/
                               df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "13")]["Nro_TK"].sum()*100))
    ibiadict[i+3]["FONIATZ"] = target[table]["FONIA"][i+3]
    ibiadict[i+3]["ADSL"] = (((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "10")]["Nro_TK_RIP_1GG"].sum() +
                               df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "10")]["Verifiche_Negative"].sum())/
                             df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "10")]["Nro_TK"].sum()*100))
    ibiadict[i+3]["ADSLTZ"] = target[table]["ADSL"][i+3]
    ibiadict[i+3]["FIBRA"] = (((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "11")]["Nro_TK_RIP_1GG"].sum() +
                                df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "11")]["Verifiche_Negative"].sum())/
                             df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "11")]["Nro_TK"].sum()*100))
    ibiadict[i+3]["FIBRATZ"] = target[table]["FIBRA"][i+3]
    ibiadict.append(dict())
    ibiadict[i+4]["date"] = (today)
    ibiadict[i+4]["group"] = (group[0].BACINO)
    ibiadict[i+4]["rowdesc"] = ("REWORK")
    ibiadict[i+4]["FONIA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "13")]["Nro_TK_RIP_1GG"].sum() /
                             df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "13")]["Nro_TK"].sum()*100))
    ibiadict[i+4]["FONIATZ"] = target[table]["FONIA"][i+4]
    ibiadict[i+4]["ADSL"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "10")]["Nro_TK_RIP_1GG"].sum() /
                             df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "10")]["Nro_TK"].sum()*100))
    ibiadict[i+4]["ADSLTZ"] = target[table]["ADSL"][i+4]
    ibiadict[i+4]["FIBRA"] = ((df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "11")]["Nro_TK_RIP_1GG"].sum() /
                             df[(df["DIAGNOSI_FE_MACRODES_id"] == "11")  & (df["Keys_SERVIZIO_id"] == "11")]["Nro_TK"].sum()*100))
    ibiadict[i+4]["FIBRATZ"] = target[table]["FIBRA"][i+4]
    return ibiadict


def sos_easy_consutivo(df, dfdays, year, month, days, tkts, discount, festdata):
    delta = 1
    paflist = dict()
    paflist["total"] = dict()
    paflist["tkts"] = list()
    totale = list()
    for i in range(0, 7):
        totale.append(0)
    for tktrow in range(0, len(list(tkts))):
        paflist["total"]["last_day"] = int(str(max(df["DATA_TK"]))[6:])
        paflist["total"]["progress"] = str(max(df["DATA_TK"]))[6:] + "/" + str(max(df["DATA_TK"]))[4:6]
        paflist["tkts"].append(dict())
        paflist["tkts"][tktrow]["service"] = tkts[tktrow].pafdesc
        paflist["tkts"][tktrow]["tktvalue"] = tkts[tktrow].value
        paflist["tkts"][tktrow]["month"] = dict()
        paflist["tkts"][tktrow]["week"] = list()
        paflist["tkts"][tktrow]["daily"] = list()
        if tkts[tktrow].pafdesc == "Tim Expert":
            paflist["tkts"][tktrow]["month"]["consuntivato_pezzi"] = df["DATA_TK"][(df["TIPO_LAVORAZIONE"].str.contains("FREE") == True)].count()
        else:
            paflist["tkts"][tktrow]["month"]["consuntivato_pezzi"] = df["DATA_TK"][(df["SERVIZIO"] == tkts[tktrow].pafdesc) & (df["TIPO_LAVORAZIONE"].str.contains("FREE") == False)].count()
        paflist["tkts"][tktrow]["month"]["consuntivato_euro"] = round(paflist["tkts"][tktrow]["month"]["consuntivato_pezzi"] * tkts[tktrow].value, 2)
        paflist["tkts"][tktrow]["month"]["preventivato_pezzi"] = 0
        paflist["tkts"][tktrow]["weektotal"] = 0
        if int(paflist["total"]["progress"][:2]) > 14:
            for i in range(0, 7):
                paflist["tkts"][tktrow]["week"].append(dict())
                paflist["tkts"][tktrow]["week"][i]["day"] = days[i].name
                if tkts[tktrow].pafdesc == "Tim Expert":
                    paflist["tkts"][tktrow]["week"][i]["avg"] = int(df["DATA_TK"][(df["ISO_WEEK_DAY"] == str(i+1)) & (df["TIPO_LAVORAZIONE"].str.contains("FREE") == True)].count() / dfdays["DATA_TK"][dfdays["ISO_WEEK_DAY"] == str(i+1)].count())
                else:
                    paflist["tkts"][tktrow]["week"][i]["avg"] = int(df["DATA_TK"][(df["ISO_WEEK_DAY"] == str(i+1)) & (df["SERVIZIO"] == str(tkts[tktrow].pafdesc)) & (df["TIPO_LAVORAZIONE"].str.contains("FREE") == False)].count() / dfdays["DATA_TK"][dfdays["ISO_WEEK_DAY"] == str(i+1)].count())
                totale[i] += paflist["tkts"][tktrow]["week"][i]["avg"]
                paflist["tkts"][tktrow]["weektotal"] += paflist["tkts"][tktrow]["week"][i]["avg"]
        else:
            if tktrow == 0:
                query = qu.dlselect_generic("ced_sos_easy_ibia", tz.sosibiadb, (str(int(year) - 1) + month[0].number))
                dfprec = pd.read_sql(query, connection)
            if dfprec.empty:
                query = qu.dlselect_generic("ced_sos_easy_ibia", tz.sosibiadb, (str(int(year)) + str(int(month[0].number)-1).zfill(2)))
                dfprec = pd.read_sql(query, connection)
                query = ("SELECT DISTINCT DATA_TK, ISO_WEEK_DAY FROM ced_sos_easy_ibia "
                         "WHERE substring(DATA_TK, 1, 6) = '%s'" % (str(int(year)) + str(int(month[0].number)-1).zfill(2)))
                dfdays = pd.read_sql(query, connection)
            else:
                if tktrow == 0:
                    query = ("SELECT DISTINCT DATA_TK, ISO_WEEK_DAY FROM ced_sos_easy_ibia "
                         "WHERE substring(DATA_TK, 1, 6) = '%s'" % (str(int(year) - 1) + month[0].number))
                    dfdays = pd.read_sql(query, connection)
                    query = qu.dlselect_generic("ced_sos_easy_ibia", tz.sosibiadb, (str(int(year)) + str(int(month[0].number)-1).zfill(2)))
                    dfmonth_1 = pd.read_sql(query, connection)
                    totdfmonth_1 = dfmonth_1["FATTURATO"][dfmonth_1["FATTURATO"] == "SI"].count()
                    query = qu.dlselect_generic("ced_sos_easy_ibia", tz.sosibiadb, (str(int(year)-1) + str(int(month[0].number)-1).zfill(2)))
                    dfprecmonth_1 = pd.read_sql(query, connection)
                    totdfprecmonth_1 = dfprecmonth_1["FATTURATO"][dfprecmonth_1["FATTURATO"] == "SI"].count()
                    delta = totdfmonth_1 / totdfprecmonth_1
            for i in range(0, 7):
                paflist["tkts"][tktrow]["week"].append(dict())
                paflist["tkts"][tktrow]["week"][i]["day"] = days[i].name
                if tkts[tktrow].pafdesc == "Tim Expert":
                    paflist["tkts"][tktrow]["week"][i]["avg"] = int(dfprec["DATA_TK"][(dfprec["ISO_WEEK_DAY"] == str(i+1)) & (dfprec["TIPO_LAVORAZIONE"].str.contains("FREE") == True)].count() / dfdays["DATA_TK"][dfdays["ISO_WEEK_DAY"] == str(i+1)].count())
                else:
                    paflist["tkts"][tktrow]["week"][i]["avg"] = int(dfprec["DATA_TK"][(dfprec["ISO_WEEK_DAY"] == str(i+1)) & (dfprec["SERVIZIO"] == str(tkts[tktrow].pafdesc)) & (dfprec["TIPO_LAVORAZIONE"].str.contains("FREE") == False)].count() / dfdays["DATA_TK"][dfdays["ISO_WEEK_DAY"] == str(i+1)].count())
                totale[i] += paflist["tkts"][tktrow]["week"][i]["avg"]
                paflist["tkts"][tktrow]["weektotal"] += paflist["tkts"][tktrow]["week"][i]["avg"]
        for i in range(1, month[0].days+1):
            paflist["tkts"][tktrow]["daily"].append(dict())
            paflist["tkts"][tktrow]["daily"][i-1]["day"] = str(i).zfill(2) + "/" + month[0].number
            paflist["tkts"][tktrow]["daily"][i-1]["nday"] = i
            try:
                paflist["tkts"][tktrow]["daily"][i-1]["wday"] = int(df["ISO_WEEK_DAY"][df["DATA_TK"].str[6:] == str(i).zfill(2)].unique())
            except:
                for day in festdata:
                    if str(day.day) == str(i):
                        paflist["tkts"][tktrow]["daily"][i-1]["wday"] = 7
                try:
                    type(paflist["tkts"][tktrow]["daily"][i-1]["wday"])
                except KeyError:
                    paflist["tkts"][tktrow]["daily"][i-1]["wday"] = dt.datetime.strptime(year + month[0].number + str(i).zfill(2), "%Y%m%d").isoweekday()
            if tkts[tktrow].pafdesc == "Tim Expert":
                paflist["tkts"][tktrow]["daily"][i-1]["ok"] = df["DATA_TK"][(df["DATA_TK"].str[6:] == str(i).zfill(2)) & (df["TIPO_LAVORAZIONE"].str.contains("FREE") == True)].count()
            else:
                paflist["tkts"][tktrow]["daily"][i-1]["ok"] = df["DATA_TK"][(df["DATA_TK"].str[6:] == str(i).zfill(2)) & (df["SERVIZIO"] == tkts[tktrow].pafdesc) & (df["TIPO_LAVORAZIONE"].str.contains("FREE") == False)].count()
            if int(paflist["tkts"][tktrow]["daily"][i-1]["day"][:2]) > int(paflist["total"]["progress"][:2]):
                paflist["tkts"][tktrow]["daily"][i-1]["ok"] = paflist["tkts"][tktrow]["week"][paflist["tkts"][tktrow]["daily"][i-1]["wday"]-1]["avg"]
                paflist["tkts"][tktrow]["month"]["preventivato_pezzi"] += paflist["tkts"][tktrow]["daily"][i-1]["ok"]
            if paflist["tkts"][tktrow]["daily"][i-1]["wday"] == 7:
                paflist["tkts"][tktrow]["daily"][i-1]["bg"] = "pink"
            else:
                paflist["tkts"][tktrow]["daily"][i-1]["bg"] = ""
        paflist["tkts"][tktrow]["month"]["preventivato_pezzi"] += paflist["tkts"][tktrow]["month"]["consuntivato_pezzi"]
        paflist["tkts"][tktrow]["month"]["preventivato_euro"] = round(decimal.Decimal(paflist["tkts"][tktrow]["month"]["preventivato_pezzi"] * tkts[tktrow].value), 2)

        ## quota quota_accantonamento per ticket (type_id == 1) (value_type == 1 number or == 2 percentual)
        tktdisc = discount[(discount["pafdesc"] == tkts[tktrow].pafdesc) & (discount["type_id"] == 1)].reset_index(drop=True)
        paflist["tkts"][tktrow]["month"]["quota_accantonamento"] = 0
        paflist["tkts"][tktrow]["month"]["sconto"] = 0
        for i in range(0, len(tktdisc.index)):
            if tktdisc["value_type_id"][i] == 1:
                paflist["tkts"][tktrow]["month"]["quota_accantonamento"] += round(paflist["tkts"][tktrow]["month"]["preventivato_pezzi"] * decimal.Decimal(tktdisc["disc_value"][i]), 2)
            else:
                paflist["tkts"][tktrow]["month"]["quota_accantonamento"] += round(paflist["tkts"][tktrow]["month"]["preventivato_pezzi"] * tkts[tktrow].value * decimal.Decimal(tktdisc["disc_value"][i]), 2)

        tktdisc = discount[(discount["pafdesc"] == tkts[tktrow].pafdesc) & (discount["type_id"] == 2)].reset_index(drop=True)
        for i in range(0, len(tktdisc.index)):
            if tktdisc["value_type_id"][i] == 1:
                paflist["tkts"][tktrow]["month"]["sconto"] += round(paflist["tkts"][tktrow]["month"]["preventivato_pezzi"] * decimal.Decimal(tktdisc["disc_value"][i]), 2)
            else:
                paflist["tkts"][tktrow]["month"]["sconto"] += round(paflist["tkts"][tktrow]["month"]["preventivato_pezzi"] * tkts[tktrow].value * decimal.Decimal(tktdisc["disc_value"][i]), 2)
        paflist["tkts"][tktrow]["month"]["preventivato_euro"] += paflist["tkts"][tktrow]["month"]["quota_accantonamento"]
    paflist["total"]["consuntivato_pezzi"] = 0
    paflist["total"]["consuntivato_euro"] = 0
    paflist["total"]["preventivato_pezzi"] = 0
    paflist["total"]["preventivato_euro"] = 0
    paflist["total"]["quota_accantonamento"] = 0
    paflist["total"]["sconto"] = 0
    grandtotal = 0
    for i in range(0, len(totale)):
        grandtotal += totale[i]
    for tktrow in range(0, len(list(tkts))):
        paflist["total"]["consuntivato_pezzi"] += paflist["tkts"][tktrow]["month"]["consuntivato_pezzi"]
        paflist["total"]["consuntivato_euro"] += paflist["tkts"][tktrow]["month"]["consuntivato_euro"]
        paflist["total"]["preventivato_pezzi"] += paflist["tkts"][tktrow]["month"]["preventivato_pezzi"]
        paflist["total"]["preventivato_euro"] += paflist["tkts"][tktrow]["month"]["preventivato_euro"]
        paflist["total"]["quota_accantonamento"] += paflist["tkts"][tktrow]["month"]["quota_accantonamento"]
        paflist["total"]["sconto"] += paflist["tkts"][tktrow]["month"]["sconto"]
    paflist["total"]["preventivato_euro"] += paflist["total"]["sconto"]
    paflist["total"]["week"] = list()
    for i in range(0, len(totale)):
        paflist["total"]["week"].append(round(totale[i]/grandtotal*100, 0))
    return paflist
