import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import datetime as dt
from string import Template
import os


def dfconvmedia(file):
    df = pd.read_excel(file, dtype={"Media_DD7": float,	"Media Fonia": float, "Media Adsl": float, "Media Fibra": float, "Media NGAN": float, "Media Fonia+Adsl": float})
    df = df.round(2)
    col_list = df.columns.values
    df.fillna(0, inplace=True)
    dfmedia = df.drop(col_list[8:], axis=1)
    dfmedia = dfmedia.drop(col_list[2:4], axis=1)
    dfmedia = dfmedia.drop(col_list[0], axis=1)
    return dfmedia


def dfconvolume(file):
    df = pd.read_excel(file, dtype="object")
    col_list = list(df.columns.values)
    df.fillna(0, inplace=True)
    dfmedia = df.drop(col_list[3:8], axis=1)
    dfmedia = dfmedia.drop(col_list[0], axis=1)
    return dfmedia


def table(df, df2, colore, i):
    check = True
    col_list1 = list(df.columns.values)
    col_list2 = list(df2.columns.values)
    strtable = """
                <table><thead><tr>
                    <th></th>
                    <th colspan='4' bgcolor='red' style='text-align: center'><font color ='white'> 187 </font></th>
                    <th></th>
                    <th colspan='3' bgcolor='red' style='text-align: center'><font color ='white'> 191 </font></th>
                </tr></thead><tbody>
                """

    for headers2 in range(0, len(col_list2)):
        while check:
            for headers1 in range(0, len(col_list1)):
                strtable += "<td bgcolor='#D3D3D3'style='text-align: center'><strong>" + col_list1[headers1] + "</strong></td>"
            strtable += "<th width='10%'> </th>"
            check =  False
        strtable += "<td bgcolor='#D3D3D3'style='text-align: center'><strong>" + col_list2[headers2] + "</strong></td>"
    for row in range(0, i):
        strtable += "<tr>"

        for data in range(0, len(col_list1)):
            bg = ""
            if (colore and data != 0 and data !=5):
                if df[str(col_list1[data])][row] >= 7.41:
                    bg = "lightgreen"
                elif df[str(col_list1[data])][row] == 0:
                    bg = "yellow"
                elif df[str(col_list1[data])][row] < 7.41:
                    bg = "pink"
                else:
                    bg = ""
            strtable += "<td style='text-align: center' bgcolor='" + bg + "'>" + str(df[str(col_list1[data])][row]) + "</td>"

        strtable += "<td width='10%'> </td>"

        for data2 in range(0, len(col_list2)):
            bg = ""
            if (colore):
                if df2[str(col_list2[data2])][row] >= 7.41:
                    bg = "lightgreen"
                elif df[str(col_list1[data])][row] == 0:
                    bg = "yellow"
                elif df2[str(col_list2[data2])][row] < 7.41:
                    bg = "pink"
                else:
                    bg = ""
            strtable += "<td style='text-align: center' bgcolor='" + bg + "'>" + str(df2[str(col_list2[data2])][row]) + "</td>"

    strtable += "</tbody></table>"
    return strtable


def tablevolume(df, df2):
    strtable = """
                <table><thead><tr>
                    <th></th>
                    <th bgcolor='red' style='text-align: center'><font color ='white'> 187 </font></th>
                    <th></th>
                    <th bgcolor='red' style='text-align: center'><font color ='white'> 191 </font></th>
                </tr></thead><tbody>
    """

    for row in range(0, len(df.index)):
        strtable += "<tr>"

        for data in range(0, 3):
            bg = ""
            if row == 0:
                bg = "bgcolor='#D3D3D3'"
            try:
                strtable += "<td style='text-align: center'" + bg + ">" + str(df[data][row]) + "</td>"
            except KeyError:
                None

        strtable += "<td width='10%'> </td>"

        for data2 in range(1, 3):
            bg = ""
            if row == 0:
                bg = "bgcolor='#D3D3D3'"
            try:
                strtable += "<td style='text-align: center' " + bg + ">" + str(df2[data2][row]) + "</td>"
            except KeyError:
                None

    strtable += "</tbody></table>"
    return strtable


def sending(table1, table2):
    PING_GOOGLE = "www.google.it"
    retcode = True
    retcode = os.system("ping -n 1 " + PING_GOOGLE)
    if retcode:
        os.system("taskkill /f /t /im ipseca.exe")

    server = "smtp.gmail.com"
    me = "Prova1234567890abcdefgh@gmail.com"
    password = "luglio2018"
    you = ["Prova1234567890abcdefgh@gmail.com", "federico.calandra@ennovagroup.it"]
#    me = "federico.calandra@ennovagroup.it"
#    password = "Consu170907!"
#    you = ["antonello.betti@ennovagroup.it", "claudia.spanu@ennovagroup.it", "federico.calandra@ennovagroup.it"]

    mesi = ("Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre")

    today = pd.read_excel(r"./excel/dataupdate.xlsx")
    today = dt.datetime.strptime(today["Data"][0], "%Y-%m-%d %H:%M")
    todaystr = dt.datetime.strftime(today, "%d/%m/%Y %H:%M")
    timeobj = dt.datetime.now()
    hourobj = timeobj.hour // 2 * 2
    timeobjstr = timeobj.replace(hour=hourobj)
    timeobjstr = timeobjstr.replace(minute=0).strftime("%d/%m/%Y %H:%M")

    htmlmail = """
    <html lang="en">
    <head>
        <style>
            table, th, td { border: 1px solid black; border-collapse: collapse; }
            th, td { padding: 5px; }
        </style>
    </head>
    <body>
    </head>
    <body><p>Ciao,</p>
    <p>di seguito i dati dei survey per il servizio C87 aggiornati il $datemailt $error</p>
    <p>Giornaliero:</p>
    {tabledayAVG}
    <p>Settimanale:</p>
    {tableweekAVG}
    <p>Mensile:</p>
    {tablemonthAVG}
    <p>Volumi risposte:</p>
    {tabledayVOLUME}
    <p>Buon lavoro,<br />Ennova Cagliari CED</p>
    </body></html>
    """

    day = dt.datetime.today()
    day = day.strftime("%d/%m/%Y")

    dfmedia1 = dfconvmedia("./excel/" + table1[-3:] + "day.xlsx")
    if len(dfmedia1.index) == 1:
        dfmedia1 = dfmedia1.append({"Data": day, "Media_DD7": 0, "Media Fonia": 0, "Media Adsl": 0, "Media Fibra": 0}, ignore_index=True)
        temp = dfmedia1.iloc[0].copy()
        dfmedia1.iloc[0] = dfmedia1.iloc[1]
        dfmedia1.iloc[1] = temp
    dfmedia2 = dfconvmedia("./excel/" + table2[-3:] + "day.xlsx")
    dfmedia2 = dfmedia2.drop(["Data", "Media Fibra"], axis=1)
    if len(dfmedia2.index) == 1:
        dfmedia2 = dfmedia2.append({"Media_DD7": 0, "Media NGAN": 0, "Media Fonia+Adsl": 0}, ignore_index=True)
        temp = dfmedia2.iloc[0].copy()
        dfmedia2.iloc[0] = dfmedia2.iloc[1]
        dfmedia2.iloc[1] = temp
    string = table(dfmedia1, dfmedia2, True, 2)
    htmlmail = htmlmail.replace("{tabledayAVG}", string)

    dfmedia1 = dfconvmedia("./excel/" + table1[-3:] + "week.xlsx")
    dfmedia2 = dfconvmedia("./excel/" + table2[-3:] + "week.xlsx")
    dfmedia2 = dfmedia2.drop(["Week", "Media Fibra"], axis=1)
    string = table(dfmedia1, dfmedia2, True, 1)
    htmlmail = htmlmail.replace("{tableweekAVG}", string)

    dfmedia1 = dfconvmedia("./excel/" + table1[-3:] + "month.xlsx")
    dfmedia1["Mese"].replace(dfmedia1["Mese"], mesi[int(dfmedia1["Mese"])-1], inplace=True)
    dfmedia2 = dfconvmedia("./excel/" + table2[-3:] + "month.xlsx")
    dfmedia2 = dfmedia2.drop(["Mese", "Media Fibra"], axis=1)
    string = table(dfmedia1, dfmedia2, True, 1)
    htmlmail = htmlmail.replace("{tablemonthAVG}", string)

    dfmedia1 = dfconvolume("./excel/" + table1[-3:] +"day.xlsx")
    dfmedia1 = dfmedia1.T
    dfmedia1 = pd.concat([dfmedia1, dfmedia1[0]], axis=1, ignore_index=True)
    dfmedia1[0] = dfmedia1.index
    dfmedia1.index = pd.RangeIndex(len(dfmedia1.index))
    if len(dfmedia1.index) > 1:
        dfmedia1 = dfmedia1.drop(1, axis=1)
    dfmedia2 = dfconvolume("./excel/" + table2[-3:] +"day.xlsx")
    dfmedia2 = dfmedia2.T
    dfmedia2 = pd.concat([dfmedia2, dfmedia2[0]], axis=1, ignore_index=True)
    dfmedia2[0] = dfmedia2.index
    dfmedia2.index = pd.RangeIndex(len(dfmedia2.index))
    if len(dfmedia2.index) > 1:
        dfmedia2 = dfmedia2.drop(1, axis=1)
    string = tablevolume(dfmedia1, dfmedia2)
    htmlmail = htmlmail.replace("{tabledayVOLUME}", string)

    errormsg = ""
    if (dt.datetime.now().hour - today.hour) > 2:
        errormsg = """
                    <br/><font color='Red'><strong>ATTENZIONE DB NON AGGIORNATI, ULTIMO AGGIORNAMENTO $datadb</font></strong>
                   """
        errormsg = Template(errormsg).safe_substitute(datadb=todaystr)

    htmlmail = Template(htmlmail).safe_substitute(datemailt=timeobjstr, error=errormsg)

    message = MIMEMultipart("alternative", None, [MIMEText(htmlmail, 'html')])

    message['Subject'] = "IVR " + timeobjstr
    message['From'] = me
    message['To'] = ", ".join(you)
    server = smtplib.SMTP_SSL(server, port=465)

    server.login(me, password)
    server.sendmail(me, you, message.as_string())
    server.quit()


# sending("c87_survey_187", "c87_survey_191")
