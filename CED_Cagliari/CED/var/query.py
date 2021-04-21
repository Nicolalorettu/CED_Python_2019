
dataivr = ("""
            SELECT IF((SELECT max(concat(str_to_date(DATA_INT,'%d/%m/%Y'),
            ' ', ORA)) FROM ced_c87_ivr_187) >(SELECT max(concat
            (str_to_date(DATA_INT,'%d/%m/%Y'),' ', ORA))
            FROM ced_c87_ivr_191),(SELECT max(concat(str_to_date
            (DATA_INT,'%d/%m/%Y'),' ', ORA)) FROM ced_c87_ivr_191),
            (SELECT max(concat(str_to_date(DATA_INT,'%d/%m/%Y'),' ', ORA))
            FROM ced_c87_ivr_187))as DATA """)


def opera_month_kpi(table, data, bacino):
    query = ("""
SELECT *
FROM """ + table + """
WHERE """ + table + """.MonthYear = '""" + data + """' AND """ + table + """.Keys_BACINO_id = '%s'
""" % bacino)
    return query

def opera_month_kpi_dl(table, data, bacino):
    query = ("""
SELECT *
FROM """ + table + """
JOIN ced_c87_group ON Keys_BACINO_id = ced_c87_group.ID
JOIN ced_opera_services ON Keys_SERVIZIO_id = ced_opera_services.ID
WHERE """ + table + """.MonthYear = '""" + data + """' AND """ + table + """.Keys_BACINO_id = '%s'
""" % bacino)
    return query



def paco_date(table, data, bacino):
    if bacino == "158":
        query = ("SELECT * FROM %s WHERE substring(Data, 4, 5)='%s'" % (table, data))
    else:
        query = ("SELECT * FROM %s WHERE substring(Data, 4, 5)='%s' AND Keys_BACINO_id = '%s'" % (table, data, bacino))
    return query


def paco_daterange(table, date_from, date_to, bacino):
    if bacino == "158":
        query = ("SELECT * FROM " + table + " WHERE str_to_date(Data,'%d-%m-%y') >= '" + date_from + "' AND str_to_date(Data,'%d-%m-%y') <= '" + date_to + "'")
    else:
        query = ("SELECT * FROM " + table + " WHERE Keys_BACINO_id = " + bacino + " AND str_to_date(Data,'%d-%m-%y') >= '" + date_from + "' AND str_to_date(Data,'%d-%m-%y') <= '" + date_to + "'")
    return query


def paco_download(table, data, bacino):
    if bacino == "158":
        query = (" SELECT * FROM %s JOIN ced_c87_group ON Keys_BACINO_id = ced_c87_group.ID JOIN ced_paco_services ON Keys_SERVIZIO_id = ced_paco_services.ID WHERE substring(Data, 4, 6)='%s'" % (table, data))
    else:
        query = (" SELECT * FROM %s JOIN ced_c87_group ON Keys_BACINO_id = ced_c87_group.ID JOIN ced_paco_services ON Keys_SERVIZIO_id = ced_paco_services.ID WHERE substring(Data, 4, 6)='%s' AND Keys_BACINO_id = '%s'" % (table, data, bacino))
    return query


def dlselect_generic(table, charn, date, group=0, z=0):
    field = list(charn)
    loops = len(charn[field[0]])
    sign = [">=", "<=", "="]
    wherecondition = ""
    middle = ""
    end = ""
    if table == "ced_sos_easy_sms_fake":
        weekcondition = ", date_format(str_to_date(substring(data_invio, 1,10), '%Y-%m-%d'), '%v') as WEEK "
    elif table == "ced_sos_easy_sms_non3xx":
        weekcondition = ", date_format(str_to_date(substring(DATA_ETT, 1,10), '%Y-%m-%d'), '%v') as WEEK "
    elif table == "ced_sos_easy_sms_risposte":
        weekcondition = ", date_format(str_to_date(substring(Data_Invio_Risposta_Cliente, 1,10), '%Y-%m-%d'), '%v') as WEEK "
    else:
        weekcondition = ""
    header = "SELECT *" + weekcondition + " FROM " + table + " WHERE "
    if isinstance(date, str):
        dump = date
        date = list()
        date.append(str(dump))
    if len(date) == 1:
        sign[0] = sign[2]
    for j in range(0, len(date)):
        if j > 0:
            wherecondition += " AND "
        for i in range(0, loops):
            if loops > 1 and i == 0:
                wherecondition += "concat("
                middle = ", "
                end = ") "
            wherecondition += ("substring(" + field[0] + ", " + charn[field[0]][i][0] + ", " + charn[field[0]][i][1] + ")")
            if i == loops - 1:
                wherecondition += end
                break
            wherecondition += middle
        wherecondition += " " + sign[j] + " '" + date[j] + "'"
    query = header + wherecondition
    if table[:9] == "ced_opera":
        query += " AND Keys_BACINO_id = '" + group[z].ID + "'"
    elif table[:7] == "ced_c87":
        if group[z].ID == "158":
            query += " AND substring(MODULO,1, 6) = '" + group[z].BACINO + "'"
        else:
            query += " AND MODULO = '" + group[z].BACINO + "'"
    elif table[:8] == "ced_paco":
        if group[z].ID == "158":
            query += ""
        else:
            query += " AND Keys_BACINO_id = '" + group[z].ID + "'"
    elif table == "ced_sos_easy_ibia":
        query += "AND tipo_intervento = 'ASSISTENZA ONLINE' AND fatturato = 'SI'"
    return query
