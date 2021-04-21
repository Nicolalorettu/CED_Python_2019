import functions.check_connect as chc
import shutil
import var.links_paths as lp
import var.sql as sql
import pandas as pd
import datetime as dt

crs = sql.createCursor()
cnx = sql.cnx


def update_kpo():
    ivrsmonth = dict()
    mdict = dict()
    gdict = dict()
    updatekpi = dict()
    dbcolumn = sql.columnkpo

    neton = chc.free_net("DataBase")
    if neton:
        import functions.dataframeoperations as dfo

        dfdict = dict()
    #    tables = ["ced_c87_ivr_187"]
        tables = ["ced_c87_ivr_187", "ced_c87_ivr_191", "ced_opera_sintesi_volumi_bo_home",
                  "ced_opera_sintesi_volumi_bo_office", "ced_opera_sintesi_volumi_fe_home",
                  "ced_opera_sintesi_volumi_fe_office", "ced_paco_richiamate_fe_187",
                  "ced_paco_richiamate_fe_191"]
        date = str(dt.datetime.now().month - 2).zfill(2) + str(dt.datetime.now().year)[2:]
        day = dt.datetime.strftime(dt.datetime.now(), "%Y-%m-%d")


        crs.execute("SELECT * FROM ced_month_tr WHERE number = '%s'" % date[:2])
        month = crs.fetchone()
        mdict[month[0]] = month[1]
        crs.execute("SELECT * FROM ced_c87_group WHERE substring(BACINO, 1, 9) = 'ENNOVA_CA' OR BACINO = 'ENNOVA'")
        groups = crs.fetchall()

        for j in range(0, len(groups)):
            group = groups[j]
            for i in range(0, len(tables)):
                if i < 2:
                    field = sql.ivrdb
                elif i > 5:
                    field = sql.pacodb
                elif i > 1:
                    field = sql.operadb
                query = sql.dlselect_generic(tables[i], field, date, group)
                if field == sql.operadb:
                    dfdict[tables[i]] = pd.read_sql(query, cnx, index_col="Keys_SERVIZIO_id")
                else:
                    dfdict[tables[i]] = pd.read_sql(query, cnx)

            updatekpi[dbcolumn[0]] = group[0] + "-" + date + "-" + day
            updatekpi[dbcolumn[1]] = date
            updatekpi[dbcolumn[2]] = day

            try:
                for i in range(0, 2):
                    ivrsmonth[tables[i]] = []
                    ivrsmonth[tables[i]].append(dfo.oldivrmonth(dfdict[tables[i]], tables[i], mdict, str(dt.datetime.now().year), 1, group[1]).copy())
                updatekpi[dbcolumn[3]] = ivrsmonth["ced_c87_ivr_187"][0][0]["Media_Fonia"]
                updatekpi[dbcolumn[4]] = ivrsmonth["ced_c87_ivr_187"][0][0]["Media_Adsl"]
                updatekpi[dbcolumn[5]] = ivrsmonth["ced_c87_ivr_187"][0][0]["Media_Fibra"]
            except:
                updatekpi[dbcolumn[3]] = 0
                updatekpi[dbcolumn[4]] = 0
                updatekpi[dbcolumn[5]] = 0

            try:
                j = 0
                colpaco = dfdict[tables[6]].columns.tolist()
                richmonth = dfo.newrichmonth(dfdict[tables[6]], tables[6], mdict, group[1], 1, colpaco)
                for i in range(6, 9):
                    updatekpi[dbcolumn[i]] = richmonth[j]["Rich_3_gg_p"]
                    j += 1
            except:
                j = 0
                colpaco = dfdict[tables[6]].columns.tolist()
                for i in range(6, 9):
                    updatekpi[dbcolumn[i]] = 0
                    j += 1

            try:
                j = 0
                colopera = dfdict[tables[4]].columns.tolist()
                kpomonth = dfo.newKpiFEHome(dfdict[tables[4]], colopera)
                for i in range(9, 21):
                    updatekpi[dbcolumn[i]] = kpomonth[j]
                    j += 1
            except:
                j = 0
                colopera = dfdict[tables[4]].columns.tolist()
                for i in range(9, 21):
                    updatekpi[dbcolumn[i]] = 0
                    j += 1

            try:
                updatekpi[dbcolumn[21]] = ivrsmonth["ced_c87_ivr_191"][0][0]["Media_Fonia_Adsl"]
                updatekpi[dbcolumn[22]] = ivrsmonth["ced_c87_ivr_191"][0][0]["Media_NGAN"]
            except:
                updatekpi[dbcolumn[21]] = 0
                updatekpi[dbcolumn[22]] = 0

            try:
                j = 0
                colpaco = dfdict[tables[7]].columns.tolist()
                richmonth = dfo.newrichmonth(dfdict[tables[7]], tables[7], mdict, group[1], 1, colpaco)
                for i in range(23, 25):
                    updatekpi[dbcolumn[i]] = richmonth[j]["Rich_3_gg_p"]
                    j += 1
            except:
                j = 0
                colpaco = dfdict[tables[7]].columns.tolist()
                for i in range(23, 25):
                    updatekpi[dbcolumn[i]] = 0
                    j += 1

            try:
                j = 0
                colopera = dfdict[tables[5]].columns.tolist()
                kpomonth = dfo.newKpiFEHome(dfdict[tables[5]], colopera)
                for i in range(25, 33):
                    updatekpi[dbcolumn[i]] = kpomonth[j]
                    j += 1
            except:
                j = 0
                colopera = dfdict[tables[5]].columns.tolist()
                for i in range(25, 33):
                    updatekpi[dbcolumn[i]] = 0
                    j += 1

            try:
                j = 0
                colopera = dfdict[tables[2]].columns.tolist()
                kpomonth = dfo.newKpiBOHoOf(dfdict[tables[2]], dfdict[tables[3]], colopera)
                for i in range(33, 44):
                    updatekpi[dbcolumn[i]] = kpomonth[j]
                    j += 1
            except:
                j = 0
                colopera = dfdict[tables[2]].columns.tolist()
                for i in range(33, 44):
                    updatekpi[dbcolumn[i]] = 0
                    j += 1

            fieldquery = ""
            valuequery = " VALUES ("
            for i in range(0, len(dbcolumn)):
                head = "REPLACE INTO ced_c87_kpi_kpo_obtained ("
                fieldquery = fieldquery + dbcolumn[i] + ", "
                middle = ")"
                valuequery = valuequery + "'" + str(updatekpi[dbcolumn[i]]) + "'" + ", "
            query = head + fieldquery[:-2] + middle + valuequery[:-2] + middle
            query = query.replace("'nan'", "'0.0'")
            crs.execute(query)
            cnx.commit()
        crs.close()
        return updatekpi
