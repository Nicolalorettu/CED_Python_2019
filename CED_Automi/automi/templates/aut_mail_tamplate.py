import decimal
import sys
import var.targets as tgt

def ivr_template_day(ivrstoday, ivrsyest, table):
    ivrsday = list()
    fieldlist = {"ced_c87_ivr_187": ["Periodo", "Media_DD7", "Media_Fonia", "Media_Adsl", "Media_Fibra"], "ced_c87_ivr_191": ["Media_DD7", "Media_NGAN", "Media_Fonia_Adsl"]}
    td = ""
    targetlist = list(tgt.newivrc87)
    htmltable = """<head>
        <style>
            table, th, td { border: 1px solid black; border-collapse: collapse; }
            th, td { padding: 5px; }
        </style>
                   </head>
                   <body>"""
    th = ("""
        <table>
            <col width="300">
            <col width="150">
            <col width="150">
            <col width="150">
            <thead>
                <tr>
                    <th bgcolor="red" style="text-align: left;vertical-align: middle" colspan="5"><font color="white"><strong><center> 187 </center></strong></font></th>
                    <td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>
                    <th bgcolor="red" style="text-align: left;vertical-align: middle" colspan="3"><font color="white"><strong><center> 191 </center></strong></font></th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td bgcolor="lightgrey"style="width:300px;text-align: center;vertical-align: middle"><strong>DATA</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>D7</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Semplici</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Medio Complesse</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Complesse</strong></td>
                    <td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>D7</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Medio Complesse</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Complesse</strong></td>
                </tr>
                """)
    ivrsday.append(ivrstoday)
    ivrsday.append(ivrsyest)

    for i in range(0, len(ivrsday)):
        td += "<tr>"
        for j in range (0, len(fieldlist["ced_c87_ivr_187"])):
            if j > 1:
                if ivrsday[i][table[0]][0][fieldlist["ced_c87_ivr_187"][j]] > tgt.newivrc87[targetlist[j-2]][2]:
                    background = "blue"
                elif ivrsday[i][table[0]][0][fieldlist["ced_c87_ivr_187"][j]] > tgt.newivrc87[targetlist[j-2]][1]:
                    background = "lightblue"
                elif ivrsday[i][table[0]][0][fieldlist["ced_c87_ivr_187"][j]] > tgt.newivrc87[targetlist[j-2]][0]:
                    background = "lightgreen"
                else:
                    background = "pink"
                td += ("""<td bgcolor='""" + background + """'style="width:300px;text-align: center;vertical-align: middle">""" + str(round(ivrsday[i][table[0]][0][fieldlist["ced_c87_ivr_187"][j]], 2)) + """</td>""")
            else:
                try:
                    td += ("""<td bgcolor=''style="width:300px;text-align: center;vertical-align: middle">""" + str(round(ivrsday[i][table[0]][0][fieldlist["ced_c87_ivr_187"][j]])) + """</td>""")
                except TypeError:
                    td += ("""<td bgcolor=''style="width:300px;text-align: center;vertical-align: middle">""" + str(ivrsday[i][table[0]][0][fieldlist["ced_c87_ivr_187"][j]]) + """</td>""")
        td += ("""<td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>""")
        for j in range (0, len(fieldlist["ced_c87_ivr_191"])):
            if j > 0:
                if ivrsday[i][table[1]][0][fieldlist["ced_c87_ivr_191"][j]] > tgt.newivrc87[targetlist[j-1]][2]:
                    background = "blue"
                elif ivrsday[i][table[1]][0][fieldlist["ced_c87_ivr_191"][j]] > tgt.newivrc87[targetlist[j-1]][1]:
                    background = "lightblue"
                elif ivrsday[i][table[1]][0][fieldlist["ced_c87_ivr_191"][j]] > tgt.newivrc87[targetlist[j-1]][0]:
                    background = "lightgreen"
                else:
                    background = "pink"
                td += ("""<td bgcolor='""" + background + """'style="width:300px;text-align: center;vertical-align: middle">""" + str(round(ivrsday[i][table[1]][0][fieldlist["ced_c87_ivr_191"][j]], 2)) + """</td>""")
            else:
                td += ("""<td bgcolor=''style="width:300px;text-align: center;vertical-align: middle">""" + str(round(ivrsday[i][table[1]][0][fieldlist["ced_c87_ivr_191"][j]], 2)) + """</td>""")
        td += "</tr>"
    htmlday = htmltable + th + td + "</table>"
    return htmlday


def ivr_template_wm(ivrs, table):
    ivrsday = list()
    fieldlist = {"ced_c87_ivr_187": ["Periodo", "Media_DD7", "Media_Fonia", "Media_Adsl", "Media_Fibra"], "ced_c87_ivr_191": ["Media_DD7", "Media_NGAN", "Media_Fonia_Adsl"]}
    td = ""
    targetlist = list(tgt.newivrc87)
    htmltable = """"""
    th = ("""
        <table>
            <col width="300">
            <col width="150">
            <col width="150">
            <col width="150">
            <thead>
                <tr>
                    <th bgcolor="red" style="text-align: left;vertical-align: middle" colspan="5"><font color="white"><strong><center> 187 </center></strong></font></th>
                    <td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>
                    <th bgcolor="red" style="text-align: left;vertical-align: middle" colspan="3"><font color="white"><strong><center> 191 </center></strong></font></th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td bgcolor="lightgrey"style="width:300px;text-align: center;vertical-align: middle"><strong>Periodo</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>D7</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Semplici</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Medio Complesse</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Complesse</strong></td>
                    <td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>D7</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Medio Complesse</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Complesse</strong></td>
                </tr>
                """)
    td += "<tr>"
    for j in range (0, len(fieldlist["ced_c87_ivr_187"])):
        if j > 1:
            if ivrs[table[0]][0][fieldlist["ced_c87_ivr_187"][j]] > tgt.newivrc87[targetlist[j-2]][2]:
                background = "blue"
            elif ivrs[table[0]][0][fieldlist["ced_c87_ivr_187"][j]] > tgt.newivrc87[targetlist[j-2]][1]:
                background = "lightblue"
            elif ivrs[table[0]][0][fieldlist["ced_c87_ivr_187"][j]] > tgt.newivrc87[targetlist[j-2]][0]:
                background = "lightgreen"
            else:
                background = "pink"
            td += ("""<td bgcolor='""" + background + """'style="width:300px;text-align: center;vertical-align: middle">""" + str(round(ivrs[table[0]][0][fieldlist["ced_c87_ivr_187"][j]], 2)) + """</td>""")
        else:
            try:
                td += ("""<td bgcolor=''style="width:300px;text-align: center;vertical-align: middle">""" + str(round(ivrs[table[0]][0][fieldlist["ced_c87_ivr_187"][j]])) + """</td>""")
            except TypeError:
                td += ("""<td bgcolor=''style="width:300px;text-align: center;vertical-align: middle">""" + str(ivrs[table[0]][0][fieldlist["ced_c87_ivr_187"][j]]) + """</td>""")
    td += ("""<td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>""")
    for j in range (0, len(fieldlist["ced_c87_ivr_191"])):
        if j > 0:
            if ivrs[table[1]][0][fieldlist["ced_c87_ivr_191"][j]] > tgt.newivrc87[targetlist[j-1]][2]:
                background = "blue"
            elif ivrs[table[1]][0][fieldlist["ced_c87_ivr_191"][j]] > tgt.newivrc87[targetlist[j-1]][1]:
                background = "lightblue"
            elif ivrs[table[1]][0][fieldlist["ced_c87_ivr_191"][j]] > tgt.newivrc87[targetlist[j-1]][0]:
                background = "lightgreen"
            else:
                background = "pink"
            td += ("""<td bgcolor='""" + background + """'style="width:300px;text-align: center;vertical-align: middle">""" + str(round(ivrs[table[1]][0][fieldlist["ced_c87_ivr_191"][j]], 2)) + """</td>""")
        else:
            td += ("""<td bgcolor=''style="width:300px;text-align: center;vertical-align: middle">""" + str(round(ivrs[table[1]][0][fieldlist["ced_c87_ivr_191"][j]], 2)) + """</td>""")
    td += "</tr>"
    htmlday = htmltable + th + td + "</table>"
    return htmlday
