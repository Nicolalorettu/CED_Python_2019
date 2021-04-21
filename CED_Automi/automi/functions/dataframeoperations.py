import pandas as pd
import datetime as dt
import numpy as np


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
    result = 0
    if type == "adsl":
        index = ["100", "103"]
    elif type == "adslnew":
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


def dfkpirich(dataframe, column, type):
    result = 0
    index = []
    if type == "semp87old":
        index = ["121", "120"]
    elif type == "semp87new":
        index = ["121", "120", "123", "124"]
    elif type == "mcmpl87old":
        index = ["122", "128", "124"]
    elif type == "mcmpl87new":
        index = ["122", "128", "126"]
    elif type == "cmpl87old":
        index = ["126", "125"]
    elif type == "cmpl187new":
        index = ["125"]
    elif type == "191old":
        index = ["130", "133", "131", "134", "132", "129"]
    elif type == "mcmpl191new":
        index = ["130", "134", "132", "133", "129"]
    elif type == "cmpl191new":
        index = ["131"]
    for i in range(0, len(index)):
        try:
            result += dataframe[column][dataframe["Keys_SERVIZIO_id"] == index[i]].sum()
        except KeyError:
            None
    return result

def oldKpiFEHome(df, columns):
    kpitable = list()
# kpitable["ReworkHF"]
    num = (df[columns[39]]["103"]+df[columns[41]]["103"])
    den = (df[columns[12]]["103"]+df[columns[33]]["103"])
    kpitable.append(num / den * 100)
# kpitable["ReworkHD"]
    num = dfkpiop(df, columns[39], "adsl")+dfkpiop(df, columns[41], "adsl")
    den = dfkpiop(df, columns[12], "adsl")+dfkpiop(df, columns[33], "adsl")
    kpitable.append(num / den * 100)
# kpitable["InvioBOHF"]
    num = (df[columns[21]]["103"]+df[columns[23]]["103"]+df[columns[34]]["103"]+df[columns[35]]["103"]
           + df[columns[36]]["103"]+df[columns[37]]["103"])
    den = (df[columns[9]]["103"]+df[columns[12]]["103"]+df[columns[21]]["103"]+df[columns[22]]["103"]+df[columns[23]]["103"]
           + df[columns[24]]["103"]+df[columns[25]]["103"]+df[columns[35]]["103"]+df[columns[36]]["103"]+df[columns[37]]["103"])
    kpitable.append(num / den * 100)
# kpitable["InvioBOHD"]
    num = (dfkpiop(df, columns[21], "adsl") + dfkpiop(df, columns[23], "adsl")
           + dfkpiop(df, columns[34], "adsl")+dfkpiop(df, columns[35], "adsl")
           + dfkpiop(df, columns[36], "adsl")+dfkpiop(df, columns[37], "adsl"))
    den = (dfkpiop(df, columns[9], "adsl")+dfkpiop(df, columns[12], "adsl")
           + dfkpiop(df, columns[21], "adsl")+dfkpiop(df, columns[22], "adsl")
           + dfkpiop(df, columns[23], "adsl")+dfkpiop(df, columns[24], "adsl")
           + dfkpiop(df, columns[25], "adsl")+dfkpiop(df, columns[35], "adsl")
           + dfkpiop(df, columns[36], "adsl")+dfkpiop(df, columns[37], "adsl"))
    kpitable.append(num / den * 100)
# kpitable["InvioOFHF"]
    num = (df[columns[24]]["103"]+df[columns[25]]["103"])
    den = (df[columns[9]]["103"]+df[columns[12]]["103"]+df[columns[21]]["103"]+df[columns[22]]["103"]+df[columns[23]]["103"]
           + df[columns[24]]["103"]+df[columns[25]]["103"]+df[columns[35]]["103"]+df[columns[36]]["103"]+df[columns[37]]["103"])
    kpitable.append(num / den * 100)
# kpitable["InvioOFHD"]
    num = dfkpiop(df, columns[24], "adsl") + dfkpiop(df, columns[25], "adsl")
    den = (dfkpiop(df, columns[9], "adsl")+dfkpiop(df, columns[12], "adsl")
           + dfkpiop(df, columns[21], "adsl")+dfkpiop(df, columns[22], "adsl")
           + dfkpiop(df, columns[23], "adsl")+dfkpiop(df, columns[24], "adsl")
           + dfkpiop(df, columns[25], "adsl")+dfkpiop(df, columns[35], "adsl")
           + dfkpiop(df, columns[36], "adsl")+dfkpiop(df, columns[37], "adsl"))
    kpitable.append(num / den * 100)
    return kpitable


def oldKpiBOHoOf(df, df2, columns):
    kpitable = list()
# kpitable["ReworkBOF"]
    num1 = (dfkpiop(df, columns[5], "bofonia") + dfkpiop(df, columns[6], "bofonia"))
    den1 = dfkpiop(df, columns[4], "bofonia")
    num2 = (dfkpiop(df2, columns[5], "bofonia") + dfkpiop(df2, columns[6], "bofonia"))
    den2 = dfkpiop(df2, columns[4], "bofonia")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["ReworkBOA"]
    num1 = (dfkpiop(df, columns[5], "boadsl") + dfkpiop(df, columns[6], "boadsl"))
    den1 = dfkpiop(df, columns[4], "boadsl")
    num2 = (dfkpiop(df2, columns[5], "boadsl") + dfkpiop(df2, columns[6], "boadsl"))
    den2 = dfkpiop(df2, columns[4], "boadsl")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["ReworkBOFi"]
    num1 = (dfkpiop(df, columns[5], "fibra") + dfkpiop(df, columns[6], "fibra"))
    den1 = dfkpiop(df, columns[4], "fibra")
    num2 = (dfkpiop(df2, columns[5], "fibra") + dfkpiop(df2, columns[6], "fibra"))
    den2 = dfkpiop(df2, columns[4], "fibra")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["InvioBOOFHF"]
    num1 = (dfkpiop(df, columns[13], "bofonia") + dfkpiop(df, columns[21], "bofonia"))
    den1 = dfkpiop(df, columns[3], "bofonia")
    num2 = (dfkpiop(df2, columns[13], "bofonia") + dfkpiop(df2, columns[21], "bofonia"))
    den2 = dfkpiop(df2, columns[3], "bofonia")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["InvioBOOFHA"]
    num1 = (dfkpiop(df, columns[13], "boadsl") + dfkpiop(df, columns[21], "boadsl"))
    den1 = dfkpiop(df, columns[3], "boadsl")
    num2 = (dfkpiop(df2, columns[13], "boadsl") + dfkpiop(df2, columns[21], "boadsl"))
    den2 = dfkpiop(df2, columns[3], "boadsl")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["InvioBOOFHFi"]
    num1 = (dfkpiop(df, columns[13], "fibra") + dfkpiop(df, columns[21], "fibra"))
    den1 = dfkpiop(df, columns[3], "fibra")
    num2 = (dfkpiop(df2, columns[13], "fibra") + dfkpiop(df2, columns[21], "fibra"))
    den2 = dfkpiop(df2, columns[3], "fibra")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["BOCTFD"]
    num1 = df[columns[11]]["100"]
    den1 = df[columns[3]]["100"]
    num2 = df2[columns[11]]["100"]
    den2 = df2[columns[3]]["100"]
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
    return kpitable


def oldKpiFEOffice(df, columns):
    kpitable = list()
# kpitable["ReworkOF"]
    num = (df[columns[39]]["100"]+df[columns[41]]["100"])
    den = (df[columns[12]]["100"]+df[columns[33]]["100"])
    kpitable.append(num / den * 100)
# kpitable["InvioBOOF"]
    num = (df[columns[21]]["100"]+df[columns[23]]["100"]+df[columns[34]]["100"]+df[columns[35]]["100"]
           + df[columns[36]]["100"]+df[columns[37]]["100"])
    den = (df[columns[9]]["100"]+df[columns[12]]["100"]+df[columns[21]]["100"]+df[columns[22]]["100"]+df[columns[23]]["100"]
           + df[columns[24]]["100"]+df[columns[25]]["100"]+df[columns[35]]["100"]+df[columns[36]]["100"]+df[columns[37]]["100"])
    kpitable.append(num / den * 100)
# kpitable["InvioOFOF"]
    num = (df[columns[24]]["100"]+df[columns[25]]["100"])
    kpitable.append(num / den * 100)
    return kpitable


def newKpiFEHome(df, columns):
    kpitable = list()
# kpitable["ReworkHF"]
    num = (df[columns[39]]["103"]+df[columns[41]]["103"])
    den = (df[columns[12]]["103"]+df[columns[35]]["103"])
    kpitable.append(num / den * 100)
# kpitable["ReworkHD"]
    num = dfkpiop(df, columns[39], "adslnew")+dfkpiop(df, columns[41], "adslnew")
    den = dfkpiop(df, columns[12], "adslnew")+dfkpiop(df, columns[35], "adslnew")
    kpitable.append(num / den * 100)
# kpitable["ReworkHFi"]
    num = dfkpiop(df, columns[39], "fibra")+dfkpiop(df, columns[41], "fibra")
    den = dfkpiop(df, columns[12], "fibra")+dfkpiop(df, columns[35], "fibra")
    kpitable.append(num / den * 100)
# kpitable["ONTHF"]
    num = df[columns[75]]["103"]
    den = df[columns[24]]["103"] + df[columns[25]]["103"]
    kpitable.append(num / den * 100)
# kpitable["ONTHD"]
    num = dfkpiop(df, columns[75], "adslnew")
    den = dfkpiop(df, columns[24], "adslnew")+dfkpiop(df, columns[25], "adslnew")
    kpitable.append(num / den * 100)
# kpitable["ONTHFi"]
    num = dfkpiop(df, columns[75], "fibra")
    den = dfkpiop(df, columns[24], "fibra")+dfkpiop(df, columns[25], "fibra")
    kpitable.append(num / den * 100)
# kpitable["InvioBOHF"]
    num = (df[columns[21]]["103"]+df[columns[23]]["103"]+df[columns[34]]["103"]+df[columns[35]]["103"]
           + df[columns[36]]["103"]+df[columns[37]]["103"])
    den = (df[columns[9]]["103"]+df[columns[12]]["103"]+df[columns[21]]["103"]+df[columns[22]]["103"]+df[columns[23]]["103"]
           + df[columns[24]]["103"]+df[columns[25]]["103"]+df[columns[35]]["103"]+df[columns[36]]["103"]+df[columns[37]]["103"])
    kpitable.append(num / den * 100)

# kpitable["InvioBOHD"]
    num = (dfkpiop(df, columns[21], "adslnew") + dfkpiop(df, columns[23], "adslnew")
           + dfkpiop(df, columns[34], "adslnew")+dfkpiop(df, columns[35], "adslnew")
           + dfkpiop(df, columns[36], "adslnew")+dfkpiop(df, columns[37], "adslnew"))
    den = (dfkpiop(df, columns[9], "adslnew")+dfkpiop(df, columns[12], "adslnew")
           + dfkpiop(df, columns[21], "adslnew")+dfkpiop(df, columns[22], "adslnew")
           + dfkpiop(df, columns[23], "adslnew")+dfkpiop(df, columns[24], "adslnew")
           + dfkpiop(df, columns[25], "adslnew")+dfkpiop(df, columns[35], "adslnew")
           + dfkpiop(df, columns[36], "adslnew")+dfkpiop(df, columns[37], "adslnew"))
    kpitable.append(num / den * 100)
# kpitable["InvioBOHFi"]
    num = (dfkpiop(df, columns[21], "fibra") + dfkpiop(df, columns[23], "fibra")
           + dfkpiop(df, columns[34], "fibra")+dfkpiop(df, columns[35], "fibra")
           + dfkpiop(df, columns[36], "fibra")+dfkpiop(df, columns[37], "fibra"))
    den = (dfkpiop(df, columns[9], "fibra")+dfkpiop(df, columns[12], "fibra")
           + dfkpiop(df, columns[21], "fibra")+dfkpiop(df, columns[22], "fibra")
           + dfkpiop(df, columns[23], "fibra")+dfkpiop(df, columns[24], "fibra")
           + dfkpiop(df, columns[25], "fibra")+dfkpiop(df, columns[35], "fibra")
           + dfkpiop(df, columns[36], "fibra")+dfkpiop(df, columns[37], "fibra"))
    kpitable.append(num / den * 100)
# kpitable["InvioOFHF"]
    num = (df[columns[24]]["103"]+df[columns[25]]["103"])
    den = (df[columns[9]]["103"]+df[columns[12]]["103"]+df[columns[21]]["103"]+df[columns[22]]["103"]+df[columns[23]]["103"]
           + df[columns[24]]["103"]+df[columns[25]]["103"]+df[columns[35]]["103"]+df[columns[36]]["103"]+df[columns[37]]["103"])
    kpitable.append(num / den * 100)
# kpitable["InvioOFHD"]
    num = dfkpiop(df, columns[24], "adslnew") + dfkpiop(df, columns[25], "adslnew")
    den = (dfkpiop(df, columns[9], "adslnew")+dfkpiop(df, columns[12], "adslnew")
           + dfkpiop(df, columns[21], "adslnew")+dfkpiop(df, columns[22], "adslnew")
           + dfkpiop(df, columns[23], "adslnew")+dfkpiop(df, columns[24], "adslnew")
           + dfkpiop(df, columns[25], "adslnew")+dfkpiop(df, columns[35], "adslnew")
           + dfkpiop(df, columns[36], "adslnew")+dfkpiop(df, columns[37], "adslnew"))
    kpitable.append(num / den * 100)
# kpitable["InvioOFHFi"]
    num = dfkpiop(df, columns[24], "fibra") + dfkpiop(df, columns[25], "fibra")
    den = (dfkpiop(df, columns[9], "fibra")+dfkpiop(df, columns[12], "fibra")
           + dfkpiop(df, columns[21], "fibra")+dfkpiop(df, columns[22], "fibra")
           + dfkpiop(df, columns[23], "fibra")+dfkpiop(df, columns[24], "fibra")
           + dfkpiop(df, columns[25], "fibra")+dfkpiop(df, columns[35], "fibra")
           + dfkpiop(df, columns[36], "fibra")+dfkpiop(df, columns[37], "fibra"))
    kpitable.append(num / den * 100)

    return kpitable


def newKpiFEOffice(df, columns):
    kpitable = list()
# kpitable["ReworkOFM"]
    num = (dfkpiop(df, columns[39], "bofonia") + dfkpiop(df, columns[41], "bofonia") +
           dfkpiop(df, columns[39], "adslnew") + dfkpiop(df, columns[41], "adslnew"))
    den = (dfkpiop(df, columns[12], "bofonia")+dfkpiop(df, columns[35], "bofonia") +
           dfkpiop(df, columns[12], "adslnew")+dfkpiop(df, columns[35], "adslnew"))
    kpitable.append(num / den * 100)
# kpitable["ReworkOFC"]
    num = dfkpiop(df, columns[39], "fibra") + dfkpiop(df, columns[41], "fibra")
    den = dfkpiop(df, columns[12], "fibra")+dfkpiop(df, columns[35], "fibra")
    kpitable.append(num / den * 100)
# kpitable["ONTMOF"]
    num = dfkpiop(df, columns[75], "bofonia") + dfkpiop(df, columns[75], "adslnew")
    den = (dfkpiop(df, columns[24], "bofonia")+dfkpiop(df, columns[25], "bofonia") +
           dfkpiop(df, columns[24], "adslnew")+dfkpiop(df, columns[25], "adslnew"))
    kpitable.append(num / den * 100)
# kpitable["ONTCOF"]
    num = dfkpiop(df, columns[75], "fibra")
    den = dfkpiop(df, columns[24], "fibra")+dfkpiop(df, columns[25], "fibra")
    kpitable.append(num / den * 100)
# kpitable["InvioBOMOF"]
    num = ((dfkpiop(df, columns[21], "bofonia") + dfkpiop(df, columns[23], "bofonia") + dfkpiop(df, columns[34], "bofonia") +
           dfkpiop(df, columns[35], "bofonia") + dfkpiop(df, columns[36], "bofonia") + dfkpiop(df, columns[37], "bofonia")) +
           (dfkpiop(df, columns[21], "adslnew") + dfkpiop(df, columns[23], "adslnew") + dfkpiop(df, columns[34], "adslnew") +
           dfkpiop(df, columns[35], "adslnew") + dfkpiop(df, columns[36], "adslnew") + dfkpiop(df, columns[37], "adslnew")))
    den = ((dfkpiop(df, columns[9], "bofonia") + dfkpiop(df, columns[12], "bofonia") + dfkpiop(df, columns[21], "bofonia") +
           dfkpiop(df, columns[22], "bofonia") + dfkpiop(df, columns[23], "bofonia") + dfkpiop(df, columns[24], "bofonia") +
           dfkpiop(df, columns[25], "bofonia") + dfkpiop(df, columns[35], "bofonia") + dfkpiop(df, columns[36], "bofonia") +
           dfkpiop(df, columns[37], "bofonia")) +
           (dfkpiop(df, columns[9], "adslnew") + dfkpiop(df, columns[12], "adslnew") + dfkpiop(df, columns[21], "adslnew") +
            dfkpiop(df, columns[22], "adslnew") + dfkpiop(df, columns[23], "adslnew") + dfkpiop(df, columns[24], "adslnew") +
            dfkpiop(df, columns[25], "adslnew") + dfkpiop(df, columns[35], "adslnew") + dfkpiop(df, columns[36], "adslnew") +
            dfkpiop(df, columns[37], "adslnew")))
    kpitable.append(num / den * 100)
# kpitable["InvioBOCOF"]
    num = (dfkpiop(df, columns[21], "fibra") + dfkpiop(df, columns[23], "fibra") + dfkpiop(df, columns[34], "fibra") +
           dfkpiop(df, columns[35], "fibra") + dfkpiop(df, columns[36], "fibra") + dfkpiop(df, columns[37], "fibra"))
    den = (dfkpiop(df, columns[9], "fibra") + dfkpiop(df, columns[12], "fibra") + dfkpiop(df, columns[21], "fibra") +
           dfkpiop(df, columns[22], "fibra") + dfkpiop(df, columns[23], "fibra") + dfkpiop(df, columns[24], "fibra") +
           dfkpiop(df, columns[25], "fibra") + dfkpiop(df, columns[35], "fibra") + dfkpiop(df, columns[36], "fibra") +
           dfkpiop(df, columns[37], "fibra"))
    kpitable.append(num / den * 100)
# kpitable["InvioOFMOF"]
    num = (dfkpiop(df, columns[24], "bofonia") + dfkpiop(df, columns[25], "bofonia") +
           dfkpiop(df, columns[24], "adslnew") + dfkpiop(df, columns[25], "adslnew"))
    den = ((dfkpiop(df, columns[9], "bofonia") + dfkpiop(df, columns[12], "bofonia") + dfkpiop(df, columns[21], "bofonia") +
           dfkpiop(df, columns[22], "bofonia") + dfkpiop(df, columns[23], "bofonia") + dfkpiop(df, columns[24], "bofonia") +
           dfkpiop(df, columns[25], "bofonia") + dfkpiop(df, columns[35], "bofonia") + dfkpiop(df, columns[36], "bofonia") +
           dfkpiop(df, columns[37], "bofonia")) +
           (dfkpiop(df, columns[9], "adslnew") + dfkpiop(df, columns[12], "adslnew") + dfkpiop(df, columns[21], "adslnew") +
            dfkpiop(df, columns[22], "adslnew") + dfkpiop(df, columns[23], "adslnew") + dfkpiop(df, columns[24], "adslnew") +
            dfkpiop(df, columns[25], "adslnew") + dfkpiop(df, columns[35], "adslnew") + dfkpiop(df, columns[36], "adslnew") +
            dfkpiop(df, columns[37], "adslnew")))
    kpitable.append(num / den * 100)
# kpitable["InvioOFCOF"]
    num = dfkpiop(df, columns[24], "fibra") + dfkpiop(df, columns[25], "fibra")
    den = (dfkpiop(df, columns[9], "fibra") + dfkpiop(df, columns[12], "fibra") + dfkpiop(df, columns[21], "fibra") +
           dfkpiop(df, columns[22], "fibra") + dfkpiop(df, columns[23], "fibra") + dfkpiop(df, columns[24], "fibra") +
           dfkpiop(df, columns[25], "fibra") + dfkpiop(df, columns[35], "fibra") + dfkpiop(df, columns[36], "fibra") +
           dfkpiop(df, columns[37], "fibra"))
    kpitable.append(num / den * 100)
    return kpitable

def newKpiBOHoOf(df, df2, columns):
    kpitable = list()
# kpitable["ReworkBOF"]
    num1 = (dfkpiop(df, columns[5], "bofonia") + dfkpiop(df, columns[6], "bofonia"))
    den1 = dfkpiop(df, columns[4], "bofonia")
    num2 = (dfkpiop(df2, columns[5], "bofonia") + dfkpiop(df2, columns[6], "bofonia"))
    den2 = dfkpiop(df2, columns[4], "bofonia")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["ReworkBOD"]
    num1 = (dfkpiop(df, columns[5], "boadsl") + dfkpiop(df, columns[6], "boadsl"))
    den1 = dfkpiop(df, columns[4], "boadsl")
    num2 = (dfkpiop(df2, columns[5], "boadsl") + dfkpiop(df2, columns[6], "boadsl"))
    den2 = dfkpiop(df2, columns[4], "boadsl")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["ReworkBOFi"]
    num1 = (dfkpiop(df, columns[5], "fibra") + dfkpiop(df, columns[6], "fibra"))
    den1 = dfkpiop(df, columns[4], "fibra")
    num2 = (dfkpiop(df2, columns[5], "fibra") + dfkpiop(df2, columns[6], "fibra"))
    den2 = dfkpiop(df2, columns[4], "fibra")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["ONTBOF"]
    num1 = dfkpiop(df, columns[27], "bofonia")
    den1 = dfkpiop(df, columns[13], "bofonia") + dfkpiop(df, columns[21], "bofonia")
    num2 = dfkpiop(df2, columns[27], "bofonia")
    den2 = dfkpiop(df2, columns[13], "bofonia") + dfkpiop(df, columns[21], "bofonia")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["ONTBOD"]
    num1 = dfkpiop(df, columns[27], "boadsl")
    den1 = dfkpiop(df, columns[13], "boadsl") + dfkpiop(df, columns[21], "boadsl")
    num2 = dfkpiop(df2, columns[27], "boadsl")
    den2 = dfkpiop(df2, columns[13], "boadsl") + dfkpiop(df, columns[21], "boadsl")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["ONTBOFi"]
    num1 = dfkpiop(df, columns[27], "fibra")
    den1 = dfkpiop(df, columns[13], "fibra") + dfkpiop(df, columns[21], "fibra")
    num2 = dfkpiop(df2, columns[27], "fibra")
    den2 = dfkpiop(df2, columns[13], "fibra") + dfkpiop(df, columns[21], "fibra")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["InvioBOOFHF"]
    num1 = (dfkpiop(df, columns[13], "bofonia") + dfkpiop(df, columns[21], "bofonia"))
    den1 = dfkpiop(df, columns[3], "bofonia")
    num2 = (dfkpiop(df2, columns[13], "bofonia") + dfkpiop(df2, columns[21], "bofonia"))
    den2 = dfkpiop(df2, columns[3], "bofonia")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["InvioBOOFHD"]
    num1 = (dfkpiop(df, columns[13], "boadsl") + dfkpiop(df, columns[21], "boadsl"))
    den1 = dfkpiop(df, columns[3], "boadsl")
    num2 = (dfkpiop(df2, columns[13], "boadsl") + dfkpiop(df2, columns[21], "boadsl"))
    den2 = dfkpiop(df2, columns[3], "boadsl")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["InvioBOOFHFi"]
    num1 = (dfkpiop(df, columns[13], "fibra") + dfkpiop(df, columns[21], "fibra"))
    den1 = dfkpiop(df, columns[3], "fibra")
    num2 = (dfkpiop(df2, columns[13], "fibra") + dfkpiop(df2, columns[21], "fibra"))
    den2 = dfkpiop(df2, columns[3], "fibra")
    kpitable.append((num1 + num2) / (den1 + den2) * 100)
# kpitable["CTEAM"] + ["RIPETIZIONE 3gg"]
    kpitable.append(0)
    kpitable.append(0)
    return kpitable


def oldivrmonth(df, table, month, year, monthcount, group):
    ivr = dict()
    ivrlist = list()
    mlist = list(month)
    segmento = "FIBRA"
    index = ["Periodo", "Gruppo", "Interviste_Valide", "Interviste_NV", "Media_DD4", "Media_DD7"]
    indexvoto = ["Voto_1", "Voto_2", "Voto_3", "Voto_4", "Voto_5", "Voto_6", "Voto_7", "Voto_8", "Voto_9", "Voto_10"]
    indextable = []
    if table == "ced_c87_ivr_187":
        indextable = ["Media_Fonia", "Media_Adsl"]
        macroesigenza = "FONIA"
    elif table == "ced_c87_ivr_191":
        indextable = ["Media_NGAN", "Media_Fonia_Adsl"]
        macroesigenza = "NGAN"
    for i in range(0, len(indextable)):
        index.append(indextable[i])
    for i in range(0, len(indexvoto)):
        index.append(indexvoto[i])
    for i in range(0, monthcount):
        ivr[index[0]] = mlist[i]
        ivr[index[1]] = group
        ivr[index[2]] = len(df[(df["DD7"] > 0) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))])
        ivr[index[3]] = len(df[(df["DD7"] == 0) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))])
        try:
            ivr[index[4]] = df[(df["DD7"] > 0) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))]["DD4"].sum() / ivr["Interviste_Valide"]
            ivr[index[5]] = df[df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year)]["DD7"].sum() / ivr["Interviste_Valide"]
        except RuntimeWarning:
            ivr[index[4]] = 0
            ivr[index[5]] = 0
        try:
            ivr[index[6]] = (df[(df["MACRO_ESIGENZA"] == macroesigenza) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))]["DD7"].sum()
                             / len(df[(df["DD7"] > 0) & (df["MACRO_ESIGENZA"] == macroesigenza) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))]))
        except RuntimeWarning:
            ivr[index[6]] = 0
        try:
            ivr[index[7]] = (df[(df["MACRO_ESIGENZA"] != macroesigenza) & (df["SEGMENTO"] != segmento) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))]["DD7"].sum()
                            / len(df[(df["DD7"] > 0) & (df["MACRO_ESIGENZA"] != macroesigenza) & (df["SEGMENTO"] != segmento) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))]))
        except RuntimeWarning:
            ivr[index[7]] = 0
        for j in range(8, 18):
            ivr[index[j]] = len(df[(df["DD7"] == j-7) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))])
        if table == "ced_c87_ivr_187":
            ivr["Media_Fibra"] = (df[(df["SEGMENTO"] == segmento) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))]["DD7"].sum()
                                  / len(df[(df["DD7"] > 0) & (df["SEGMENTO"] == segmento) & (df.DATA_INT.str[3:] == (month[mlist[i]] + "/" + year))]))
        ivrlist.append(ivr.copy())
    return ivrlist


def oldrichmonth(df, table, month, group, monthcount, columns):
    rich = dict()
    richlist = list()
    monthlist = list(month)
    indextable = []
    index = ["Periodo", "Gruppo", "Tipo", "Chiamate", "Abbattute", "Rich. 3 gg", "Rich. 3 gg %"]
    if table == "ced_paco_richiamate_fe_187":
        indextable = ["semp87old", "mcmpl87old", "cmpl87old"]
    elif table == "ced_paco_richiamate_fe_191":
        indextable = ["191old"]
    for i in range(0, monthcount):                 # Conto Mese Totale
        for j in range(0, len(indextable)):                   # Conto Mese Tipo
            rich[index[0]] = monthlist[i]
            rich[index[1]] = group
            rich[index[2]] = indextable[j]
            rich[index[3]] = dfkpirich(df, columns[4], indextable[j])
            rich[index[4]] = dfkpirich(df, columns[5], indextable[j])
            rich[index[5]] = dfkpirich(df, columns[8], indextable[j])
            try:
                rich[index[6]] = rich[index[5]] / rich[index[3]]*100
            except RuntimeWarning:
                rich[index[6]] = 0
            richlist.append(rich.copy())
    return richlist


def newrichmonth(df, table, month, group, monthcount, columns):
    rich = dict()
    richlist = list()
    monthlist = list(month)
    indextable = []
    index = ["Periodo", "Gruppo", "Tipo", "Chiamate", "Abbattute", "Rich. 3 gg", "Rich. 3 gg %"]
    if table == "ced_paco_richiamate_fe_187":
        indextable = ["semp87new", "mcmpl87new", "cmpl187new"]
    elif table == "ced_paco_richiamate_fe_191":
        indextable = ["mcmpl191new", "cmpl191new"]
    for i in range(0, monthcount):                 # Conto Mese Totale
        for j in range(0, len(indextable)):                   # Conto Mese Tipo
            rich[index[0]] = monthlist[i]
            rich[index[1]] = group
            rich[index[2]] = indextable[j]
            rich[index[3]] = dfkpirich(df, columns[4], indextable[j])
            rich[index[4]] = dfkpirich(df, columns[5], indextable[j])
            rich[index[5]] = dfkpirich(df, columns[8], indextable[j])
            try:
                rich[index[6]] = rich[index[5]] / rich[index[3]]*100
            except RuntimeWarning:
                rich[index[6]] = 0
            richlist.append(rich.copy())
    return richlist


def oldrichweek(df, table, fweek, lweek, group, columns):
    rich = dict()
    richlist = list()
    indextable = []
    index = ["Periodo", "Gruppo", "Tipo", "Chiamate", "Abbattute", "Rich. 3 gg", "Rich. 3 gg %"]
    try:
        df["Data"] = pd.to_datetime(df["Data"], dayfirst=True)
        df["Data"] = df["Data"].dt.strftime("%W")
    except ValueError:
        None
    if table == "ced_paco_richiamate_fe_187":
        indextable = ["semp87old", "mcmpl87old", "cmpl87old"]
    elif table == "ced_paco_richiamate_fe_191":
        indextable = ["191old"]
    for week in range(fweek, lweek+1):
        for i in range(0, len(indextable)):
            rich[index[0]] = week
            rich[index[1]] = group
            rich[index[2]] = indextable[i]
            rich[index[3]] = dfkpirich(df[df.Data == str(week)], columns[4], indextable[i])
            rich[index[4]] = dfkpirich(df[df.Data == str(week)], columns[5], indextable[i])
            rich[index[5]] = dfkpirich(df[df.Data == str(week)], columns[8], indextable[i])
            try:
                rich[index[6]] = rich[index[5]] / rich[index[3]]*100
            except RuntimeWarning:
                rich[index[6]] = 0
            richlist.append(rich.copy())
    return richlist


def newrichweek(df, table, fweek, lweek, group, columns):
    rich = dict()
    richlist = list()
    indextable = []
    index = ["Periodo", "Gruppo", "Tipo", "Chiamate", "Abbattute", "Rich. 3 gg", "Rich. 3 gg %"]
    try:
        df["Data"] = pd.to_datetime(df["Data"], dayfirst=True)
        df["Data"] = df["Data"].dt.strftime("%W")
    except ValueError:
        None
    if table == "ced_paco_richiamate_fe_187":
        indextable = ["semp87new", "mcmpl87new", "cmpl187new"]
    elif table == "ced_paco_richiamate_fe_191":
        indextable = ["mcmpl191new", "cmpl191new"]
    for week in range(fweek, lweek+1):
        for i in range(0, len(indextable)):
            rich[index[0]] = week
            rich[index[1]] = group
            rich[index[2]] = indextable[i]
            rich[index[3]] = dfkpirich(df[df.Data == str(week)], columns[4], indextable[i])
            rich[index[4]] = dfkpirich(df[df.Data == str(week)], columns[5], indextable[i])
            rich[index[5]] = dfkpirich(df[df.Data == str(week)], columns[8], indextable[i])
            try:
                rich[index[6]] = rich[index[5]] / rich[index[3]]*100
            except RuntimeWarning:
                rich[index[6]] = 0
            richlist.append(rich.copy())
    return richlist


def oldivrweek(df, table, fweek, lweek, group):
    np.seterr(divide='ignore', invalid='ignore')
    ivr = dict()
    ivrlist = list()
    segmento = "FIBRA"
    index = ["Periodo", "Gruppo", "Interviste_Valide", "Interviste_NV", "Media_DD4", "Media_DD7"]
    indexvoto = ["Voto_1", "Voto_2", "Voto_3", "Voto_4", "Voto_5", "Voto_6", "Voto_7", "Voto_8", "Voto_9", "Voto_10"]
    indextable = []
    try:
        df["DATA_INT"] = pd.to_datetime(df["DATA_INT"], dayfirst=True)
        df["DATA_INT"] = df["DATA_INT"].dt.strftime("%W")
    except ValueError:
        None
    if table == "ced_c87_ivr_187":
        indextable = ["Media_Fonia", "Media_Adsl"]
        macroesigenza = "FONIA"
    elif table == "ced_c87_ivr_191":
        indextable = ["Media_NGAN", "Media_Fonia_Adsl"]
        macroesigenza = "NGAN"
    for i in range(0, len(indextable)):
        index.append(indextable[i])
    for i in range(0, len(indexvoto)):
        index.append(indexvoto[i])
    for week in range(fweek, lweek+1):
        ivr[index[0]] = week
        ivr[index[1]] = group
        ivr[index[2]] = len(df[(df["DD7"] > 0) & (df.DATA_INT == str(week))])
        ivr[index[3]] = len(df[(df["DD7"] == 0) & (df.DATA_INT == str(week))])
        ivr[index[4]] = df[(df["DD7"] > 0) & (df.DATA_INT == str(week))]["DD4"].sum() / ivr["Interviste_Valide"]
        ivr[index[5]] = df[df.DATA_INT == str(week)]["DD7"].sum() / ivr["Interviste_Valide"]
        ivr[index[6]] = (df[(df["MACRO_ESIGENZA"] == macroesigenza) & (df.DATA_INT == str(week))]["DD7"].sum()
                        / len(df[(df["DD7"] > 0) & (df["MACRO_ESIGENZA"] == macroesigenza) & (df.DATA_INT == str(week))]))
        ivr[index[7]] = (df[(df["MACRO_ESIGENZA"] != macroesigenza) & (df["SEGMENTO"] != segmento) & (df.DATA_INT == str(week))]["DD7"].sum()
                        / len(df[(df["DD7"] > 0) & (df["MACRO_ESIGENZA"] != macroesigenza) & (df["SEGMENTO"] != segmento) & (df.DATA_INT == str(week))]))
        for j in range(8, 18):
            ivr[index[j]] = len(df[(df["DD7"] == j-7) & (df.DATA_INT == str(week))])
        if table == "ced_c87_ivr_187":
            ivr["Media_Fibra"] = (df[(df["SEGMENTO"] == segmento)  & (df.DATA_INT == str(week))]["DD7"].sum()
                                 / len(df[(df["DD7"] > 0) & (df["SEGMENTO"] == segmento) & (df.DATA_INT == str(week))]))
        ivrlist.append(ivr.copy())
    return ivrlist
