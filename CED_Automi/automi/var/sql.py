import mysql.connector as mc
import datetime as dt
#pychrome

# DB DATE FIELDS CHARACTER FOR MONTH

operadb = {"MonthYear": [["1", "4"]]}
ivrdb = {"DATA_INT": [["4", "2"], ["9", "2"]]}
pacodb = {"Data": [["4", "2"], ["7", "2"]]}

columnkpo = ["ID", "MonthYear", "Day", "FEH_Ivr_Semplici", "FEH_Ivr_Medio_Complesse",
            "FEH_Ivr_Home_Complesse", "FEH_Richiamate_Semplici", "FEH_Richiamate_Medio_Complesse",
            "FEH_Richiamate_Complesse", "FEH_Rework_FONIA_HOME", "FEH_Rework_DATI_HOME",
            "FEH_Rework_FIBRA_HOME", "FEH_ONT_Fonia", "FEH_ONT_ADSL", "FEH_ONT_Fibra",
            "FEH_Invio_BO_Fonia", "FEH_Invio_BO_ADSL", "FEH_Invio_BO_Fibra",
            "FEH_Invio_OF_Fonia", "FEH_Invio_OF_ADSL", "FEH_Invio_OF_Fibra",
            "FEO_Ivr_Medio_Complesse", "FEO_Ivr_Complesse", "FEO_Richiamate_Medio_Complesse",
            "FEO_Richiamate_Complesse", "FEO_Rework_Medie", "FEO_Rework_Complesse",
            "FEO_ONT_Medie", "FEO_ONT_Complesse", "FEO_Invio_BO_Medie",
            "FEO_Invio_BO_Complesse", "FEO_Invio_OF_Medie", "FEO_Invio_OF_Complesse",
            "BO_BO_REWORK_VN_RIP7GG_FONIA", "BO_BO_REWORK_VN_RIP7GG_ADSL",
            "BO_BO_REWORK_VN_RIP7GG_FIBRA", "BO_ONT_Fonia", "BO_ONT_ADSL", "BO_ONT_Fibra",
            "BO_BO_INVIO_OF_FONIA", "BO_BO_INVIO_OF_ADSL", "BO_BO_INVIO_OF_FIBRA",
            "BO_C_TEAM_FONIA_DATI", "BO_RIPETIZIONE_A_33_GG_SU_COLLAUDI"]


# DB PARAMETER, CURSOR AND QUERIES

cnx = mc.connect(user='root',
                 password='Levissima2018!',
                 host='127.0.0.1',
                 database='ced_cagliari')


def createCursor():
    return cnx.cursor(buffered=True)


# REPLACE IN SQL WITH CSV FILES
def repwithcsv(path, table, lines):
    query = ("LOAD DATA INFILE '" + path + "' REPLACE INTO TABLE " + table +
             " FIELDS TERMINATED BY ';' IGNORE " + lines + " LINES")
    return query


def cred_query(tool):
    query = ("SELECT * FROM ced_tools_credentials WHERE description = '%s'" % tool)
    return query


def upddbtracker(service):
    query = ("INSERT INTO ced_db_update_tracker (`db_name`, `date_hour`) VALUES ( '%s', '%s')" % (service, dt.datetime.now().strftime("%d/%m/%Y %H:%M")))
    return query

def dlselect_generic(table, charn, date, group):
    field = list(charn)
    loops = len(charn[field[0]])
    wherecondition = ""
    middle = ""
    end = ""
    header = "SELECT * FROM " + table + " WHERE "
    for i in range(0, loops):
        if loops > 1 and i == 0:
            wherecondition = "concat("
            middle = ", "
            end = ") "
        wherecondition += ("substring(" + field[0] + ", " + charn[field[0]][i][0] + ", " + charn[field[0]][i][1] + ")")
        if i == loops - 1:
            wherecondition += end
            break
        wherecondition += middle
    query = header + wherecondition + "= '" + date + "'"
    if table[:9] == "ced_opera":
        query += " AND Keys_BACINO_id = '" + group[0] + "'"
    elif table[:7] == "ced_c87":
        if group[0] == "158":
            query += " AND substring(MODULO,1, 6) = '" + group[1] + "'"
        else:
            query += " AND MODULO = '" + group[1] + "'"
    elif table[:8] == "ced_paco":
        if group[0] == "158":
            query += ""
        else:
            query += " AND Keys_BACINO_id = '" + group[0] + "'"
    return query

def month_query(month):
    query = ("SELECT * FROM ced_month_tr WHERE number = '%s'" % month)
    return query

def month_query_name(month):
    query = ("SELECT * FROM ced_month_tr WHERE name = '%s'" % month)
    return query


def insert_query(table, df, j):
    columns = df.columns
    query = "INSERT INTO " + table + " ("
    for i in range(0, len(columns)):
        if i != 0:
            query += ", "
        query += columns[i]
    query += ") VALUES ("
    for i in range(0, len(columns)):
        if i != 0:
            query += ", "
        try:
            query += "'" + str(df[columns[i]][j]) + "'"
        except KeyError:
            query += "'0'"
    query += ")"
    return query


def update_SMS(table, value):
    query = "UPDATE " + table + " SET PRIMO_SMS = 'SI' WHERE Cod_Risposta = '" + value + "'"
    return query
