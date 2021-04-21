import decimal
import sys

def kpo_template(kpodict, table, form, group, z):
    td = dict()
    htmltable = """<head>
        <style>
            table, th, td { border: 1px solid black; border-collapse: collapse; }
            th, td { padding: 5px; }
        </style>
                   </head>
                   <body>"""
    thfeh = ("""
        <table>
            <col width="300">
            <col width="150">
            <col width="150">
            <col width="150">
            <thead>
                <tr>
                    <th bgcolor="red" style="text-align: left;vertical-align: middle" colspan="4"><font color="white">Report KPI FE - """
                    + form["month"] + """/""" + form["year"] + """</font></th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td bgcolor="lightgrey"style="width:300px;text-align: center;vertical-align: middle"><strong>INBOUND HOME</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>KPI</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>""" + group[z].BACINO + """</strong></td>
                    <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Delta vs KPI</strong></td>
                </tr>
                """)
    for i in range(0, len(table)):
        td[table[i]] = list()
        j = 0
        for kpi in kpodict[table[i]]:
            td[table[i]].append("")
            td[table[i]][j] = ("""
                    <tr>
                        <td bgcolor=""style="text-align: left;vertical-align: middle">""" + kpi["name"] + """</td> """)
            if kpi["kpi_type"] == 1:
                if kpi["target"] != 0:
                    td[table[i]][j] += ("""
                            <td bgcolor=""style="text-align: center;vertical-align: middle">""" + str(round(kpi["target"], 2)) + """</td>""")
                else:
                    td[table[i]][j] += ("""
                            <td bgcolor=""style="text-align: center;vertical-align: middle">""" + str(round(kpi["tier1"], 2)) + """</td>""")
                if kpi["delta"] > 0:
                    bgcolor = '"lightgreen"'
                else:
                    bgcolor = '"pink"'
                td[table[i]][j] += ("""
                        <td bgcolor=""" + bgcolor + """style="text-align: center;vertical-align: middle">""" + str(round(kpi["kpo"], 2)) + """</td>
                        <td bgcolor=""style="text-align: center;vertical-align: middle">""" + str(round(kpi["delta"], 2)) + """</td>
                    </tr>
                        """)
            else:
                if kpi["target"] != 0:
                    td[table[i]][j] += ("""
                            <td bgcolor=""style="text-align: center;vertical-align: middle">""" + str(round(kpi["target"], 2)) + """ %</td>""")
                else:
                    td[table[i]][j] += ("""
                            <td bgcolor=""style="text-align: center;vertical-align: middle">""" + str(round(kpi["tier1"], 2)) + """ %</td>""")
                if kpi["delta"] > 0:
                    bgcolor = '"lightgreen"'
                else:
                    bgcolor = '"pink"'
                td[table[i]][j] += ("""
                        <td bgcolor=""" + bgcolor + """style="text-align: center;vertical-align: middle">""" + str(round(kpi["kpo"], 2)) + """ %</td>
                        <td bgcolor=""style="text-align: center;vertical-align: middle">""" + str(round(kpi["delta"], 2)) + """ %</td>
                    </tr>
                        """)
            j += 1
    septable = """
            <tr>
                <td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>
                <td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>
                <td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>
                <td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>
                <td bgcolor=""style="text-align: center;vertical-align: middle; border-color:white"></td>
            </tr>
            """
    thfeo = ("""
            <tr>
                <td bgcolor="lightgrey"style="width:300px;text-align: center;vertical-align: middle"><strong>INBOUND OFFICE</strong></td>
                <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>KPI</strong></td>
                <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>""" + group[z].BACINO + """</strong></td>
                <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Delta vs KPI</strong></td>
            </tr>
            """)
    endtb = """
                </tbody>
            </table>
        """
    endfile = """
                </tbody>
            </table>
    </body>
        """
    thbo = ("""
                <table>
                    <col width="300">
                    <col width="150">
                    <col width="150">
                    <col width="150">
                    <thead>
                        <tr>
                            <th bgcolor="red"style="text-align: left;vertical-align: middle" colspan="4"><font color="white"><strong>Report KPI - """
                            + form["month"] + """/""" + form["year"] + """
                            </font></th>
                    </thead>
                    <tbody>
                        <tr>
                            <td bgcolor="lightgrey"style="width:300px;text-align: center;vertical-align: middle"><strong>HOME+OFFICE</strong></td>
                            <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>KPI</strong></td>
                            <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>""" + group[z].BACINO + """</strong></td>
                            <td bgcolor="lightgrey"style="width:150px;text-align: center;vertical-align: middle"><strong>Delta vs KPI</strong></td>
                        </tr>
            """)

    for i in range(0, len(table)):
        if i == 0:
            htmltable += thfeh
        elif i == 1:
            htmltable += thfeo
        elif i == 2:
            htmltable += thbo
        for j in range(0, len(td[table[i]])):
            htmltable += td[table[i]][j]
        htmltable += septable
        if i == 1:
            htmltable += endtb
    htmltable += endtb
    return htmltable


def header(service, month):
    html = """
    <p>Ciao,</p>
<p> </p>
    <p>In allegato il report """ + service + """ riferito al mese di """ + month + """</p>
    """
    return html


def footer():
    html = """
    <p>Buon Lavoro,</p>
    <p>Portale CED Cagliari</p>
    """
    return html
