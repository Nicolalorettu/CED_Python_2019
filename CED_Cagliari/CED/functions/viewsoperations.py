try:
    import dataframeoperations as dfo
except:
    import CED.functions.dataframeoperations as dfo
import datetime as dt
import pandas as pd
from django.db import connection
from CED.models import c87_group

def c872018_kpi(dfdict, tables, mdict, group, grcount, target, typeupd, lastday, date, dateivr, contract):
    try:
        ivrsmonth = dict()
        updatekpi = dict()
        updatekpi["FEH"] = list()
        updatekpi["FEO"] = list()
        updatekpi["BO"] = list()
        date = mdict
        for i in range(0, 2):
            ivrsmonth[tables[i]] = []
            ivrsmonth[tables[i]].append(dfo.ivrmonth(dfdict[tables[i]], tables[i], dateivr, group[grcount].ID).copy())
        updatekpi["FEH"].append(dict())
        updatekpi["FEH"][0]["name"] = target[0].name
        updatekpi["FEH"][0]["target"] = target[0].target
        updatekpi["FEH"][0]["tier1"] = target[0].tier1
        updatekpi["FEH"][0]["tier2"] = target[0].tier2
        updatekpi["FEH"][0]["tier3"] = target[0].tier3
        updatekpi["FEH"][0]["kpo"] = ivrsmonth["ced_c87_ivr_187"][0][0]["Media_Fonia"]
        updatekpi["FEH"].append(dict())
        updatekpi["FEH"][1]["name"] = target[1].name
        updatekpi["FEH"][1]["target"] = target[1].target
        updatekpi["FEH"][1]["tier1"] = target[1].tier1
        updatekpi["FEH"][1]["tier2"] = target[1].tier2
        updatekpi["FEH"][1]["tier3"] = target[1].tier3
        updatekpi["FEH"][1]["kpo"] = ivrsmonth["ced_c87_ivr_187"][0][0]["Media_Adsl"]
        updatekpi["FEH"].append(dict())
        updatekpi["FEH"][2]["name"] = target[2].name
        updatekpi["FEH"][2]["target"] = target[2].target
        updatekpi["FEH"][2]["tier1"] = target[2].tier1
        updatekpi["FEH"][2]["tier2"] = target[2].tier2
        updatekpi["FEH"][2]["tier3"] = target[2].tier3
        updatekpi["FEH"][2]["kpo"] = ivrsmonth["ced_c87_ivr_187"][0][0]["Media_Fibra"]
        j = 0
        colpaco = dfdict[tables[6]].columns.tolist()
        richmonth = dfo.richmonth(dfdict[tables[6]], tables[6], mdict, group[grcount].BACINO, colpaco, contract)
        for i in range(3, 6):
            updatekpi["FEH"].append(dict())
            updatekpi["FEH"][i]["name"] = target[i].name
            updatekpi["FEH"][i]["target"] = target[i].target
            updatekpi["FEH"][i]["tier1"] = target[i].tier1
            updatekpi["FEH"][i]["tier2"] = target[i].tier2
            updatekpi["FEH"][i]["tier3"] = target[i].tier3
            updatekpi["FEH"][i]["kpo"] = richmonth[j]["Rich_3_gg_p"]
            j += 1
        j = 0
        colopera = dfdict[tables[4]].columns.tolist()
        kpomonth = dfo.KpiFEHome2018(dfdict[tables[4]], colopera)
        for i in range(6, 18):
            updatekpi["FEH"].append(dict())
            updatekpi["FEH"][i]["name"] = target[i].name
            updatekpi["FEH"][i]["target"] = target[i].target
            updatekpi["FEH"][i]["tier1"] = target[i].tier1
            updatekpi["FEH"][i]["tier2"] = target[i].tier2
            updatekpi["FEH"][i]["tier3"] = target[i].tier3
            updatekpi["FEH"][i]["kpo"] = kpomonth[j]
            j += 1
        updatekpi["FEO"].append(dict())
        updatekpi["FEO"][0]["name"] = target[18].name
        updatekpi["FEO"][0]["target"] = target[18].target
        updatekpi["FEO"][0]["tier1"] = target[18].tier1
        updatekpi["FEO"][0]["tier2"] = target[18].tier2
        updatekpi["FEO"][0]["tier3"] = target[18].tier3
        updatekpi["FEO"][0]["kpo"] = ivrsmonth["ced_c87_ivr_191"][0][0]["Media_Fonia_Adsl"]
        updatekpi["FEO"].append(dict())
        updatekpi["FEO"][1]["name"] =  target[19].name
        updatekpi["FEO"][1]["target"] = target[19].target
        updatekpi["FEO"][1]["tier1"] = target[19].tier1
        updatekpi["FEO"][1]["tier2"] = target[19].tier2
        updatekpi["FEO"][1]["tier3"] = target[19].tier3
        updatekpi["FEO"][1]["kpo"] = ivrsmonth["ced_c87_ivr_191"][0][0]["Media_NGAN"]
        j = 0
        colpaco = dfdict[tables[7]].columns.tolist()
        richmonth = dfo.richmonth(dfdict[tables[7]], tables[7], mdict, group[grcount].BACINO, colpaco, contract)
        for i in range(2, 4):
            updatekpi["FEO"].append(dict())
            updatekpi["FEO"][i]["name"] =  target[18+i].name
            updatekpi["FEO"][i]["target"] = target[18+i].target
            updatekpi["FEO"][i]["tier1"] = target[18+i].tier1
            updatekpi["FEO"][i]["tier2"] = target[18+i].tier2
            updatekpi["FEO"][i]["tier3"] = target[18+i].tier3
            updatekpi["FEO"][i]["kpo"] = richmonth[j]["Rich_3_gg_p"]
            j += 1
        j = 0
        colopera = dfdict[tables[5]].columns.tolist()
        kpomonth = dfo.KpiFEHome2018(dfdict[tables[5]], colopera)
        for i in range(4, 12):
            updatekpi["FEO"].append(dict())
            updatekpi["FEO"][i]["name"] =  target[18+i].name
            updatekpi["FEO"][i]["target"] = target[18+i].target
            updatekpi["FEO"][i]["tier1"] = target[18+i].tier1
            updatekpi["FEO"][i]["tier2"] = target[18+i].tier2
            updatekpi["FEO"][i]["tier3"] = target[18+i].tier3
            updatekpi["FEO"][i]["kpo"] = kpomonth[j]
            j += 1
        j = 0
        colopera = dfdict[tables[2]].columns.tolist()
        kpomonth = dfo.KpiBOHoOf2018(dfdict[tables[2]], dfdict[tables[3]], colopera)
        for i in range(0, 11):
            updatekpi["BO"].append(dict())
            updatekpi["BO"][i]["name"] = target[30+i].name
            updatekpi["BO"][i]["target"] = target[30+i].target
            updatekpi["BO"][i]["tier1"] = target[30+i].tier1
            updatekpi["BO"][i]["tier2"] = target[30+i].tier2
            updatekpi["BO"][i]["tier3"] = target[30+i].tier3
            updatekpi["BO"][i]["kpo"] = kpomonth[j]
            j += 1
        k = 0
        for i in typeupd:
            for j in range(0, len(updatekpi[i])):
                if updatekpi[i][j]["kpo"] != 0:
                    if target[j+k].kpi_type.id == 1:
                        updatekpi[i][j]["delta"] = updatekpi[i][j]["kpo"] - updatekpi[i][j]["target"]
                    elif target[j+k].kpi_type.id == 2:
                        updatekpi[i][j]["delta"] = updatekpi[i][j]["target"] - updatekpi[i][j]["kpo"]
                else:
                    updatekpi[i][j]["delta"] = 0
                updatekpi[i][j]["kpi_type"] = target[j+k].kpi_type.id
            k += j+1

        today = dt.datetime.now().date()
        x = 1
        try:
            if lastday[0].Day != today:
                dflastkpo = pd.read_sql("SELECT * FROM ced_c87_kpi_kpo_obtained WHERE MonthYear = '%s' AND substring(id,1,3) = '%s' AND Day = '%s'"
                                        % (date, group[grcount].ID, dt.datetime.strftime(lastday[0].Day, "%Y-%m-%d")), connection)
            else:
                dflastkpo = pd.read_sql("SELECT * FROM ced_c87_kpi_kpo_obtained WHERE MonthYear = '%s' AND "
                                        "substring(id,1,3) = '%s' AND Day < '%s' AND Day >= '%s'"
                                        % (date, group[grcount].ID, dt.datetime.strftime(today, "%Y-%m-%d"),
                                            dt.datetime.strftime(today - dt.timedelta(x), "%Y-%m-%d")),
                                        connection)
        except:
            dflastkpo = 0
        z = 3
        if dflastkpo == 0:
            for i in range(0, len(updatekpi)):
                for j in range(0, len(updatekpi[typeupd[i]])):
                    updatekpi[typeupd[i]][j]["old_kpo"] = 0
        else:
            collastkpo = dflastkpo.columns.tolist()
            for i in range(0, len(updatekpi)):
                for j in range(0, len(updatekpi[typeupd[i]])):
                    updatekpi[typeupd[i]][j]["old_kpo"] = dflastkpo[collastkpo[z]][0]
                    z += 1

        z = 0
        for i in range(0, len(updatekpi)):
            for j in range(0, len(updatekpi[typeupd[i]])):
                if target[z].kpi_type_id == 1:
                    updatekpi[typeupd[i]][j]["malus_target"] = 0
                    updatekpi[typeupd[i]][j]["malus_tier1"] = updatekpi[typeupd[i]][j]["kpo"] - target[z].tier1
                    if updatekpi[typeupd[i]][j]["kpo"] < target[z].tier1:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier1
                    elif updatekpi[typeupd[i]][j]["kpo"] < target[z].tier2:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier2
                    elif updatekpi[typeupd[i]][j]["kpo"] < target[z].tier3:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier3
                    else:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier4

                elif target[z].kpi_type_id == 2:

                    if target[z].target > 0:
                        if updatekpi[typeupd[i]][j]["kpo"] > target[z].target:
                            updatekpi[typeupd[i]][j]["malus_target"] = target[z].target - updatekpi[typeupd[i]][j]["kpo"]
                        else:
                            updatekpi[typeupd[i]][j]["malus_target"] = 0
                        if updatekpi[typeupd[i]][j]["kpo"] > target[z].tier1:
                            updatekpi[typeupd[i]][j]["malus_tier1"] = target[z].tier1 - updatekpi[typeupd[i]][j]["kpo"] - updatekpi[typeupd[i]][j]["malus_target"]
                            updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier1
                        elif updatekpi[typeupd[i]][j]["kpo"] > target[z].tier2:
                            updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier2
                        elif updatekpi[typeupd[i]][j]["kpo"] > target[z].tier3:
                            updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier3
                        else:
                            updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier4
                    else:
                        updatekpi[typeupd[i]][j]["malus_target"] = 0
                        if updatekpi[typeupd[i]][j]["kpo"] > target[z].tier1:
                            updatekpi[typeupd[i]][j]["malus_tier1"] = target[z].tier1 - updatekpi[typeupd[i]][j]["kpo"]
                            updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier1
                        elif updatekpi[typeupd[i]][j]["kpo"] > target[z].tier2:
                            updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier2
                        elif updatekpi[typeupd[i]][j]["kpo"] > target[z].tier3:
                            updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier3
                        else:
                            updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier4
                z += 1
    except Exception as error:
        print(type(error), error)
        updatekpi = {}
    return updatekpi


def c872017_kpi(dfdict, tables, mdict, group, grcount, target, typeupd, lastday, date, dateivr, contract):
    try:
        ivrsmonth = dict()
        updatekpi = dict()
        updatekpi["FEH"] = list()
        updatekpi["FEO"] = list()
        updatekpi["BO"] = list()
        for i in range(0, 2):
            ivrsmonth[tables[i]] = []
            ivrsmonth[tables[i]].append(dfo.ivrmonth(dfdict[tables[i]], tables[i], dateivr, group[grcount].ID).copy())
        updatekpi["FEH"].append(dict())
        updatekpi["FEH"][0]["name"] = target[0].name
        updatekpi["FEH"][0]["target"] = target[0].target
        updatekpi["FEH"][0]["tier1"] = target[0].tier1
        updatekpi["FEH"][0]["tier2"] = target[0].tier2
        updatekpi["FEH"][0]["tier3"] = target[0].tier3
        updatekpi["FEH"][0]["kpo"] = ivrsmonth["ced_c87_ivr_187"][0][0]["Media_DD7"]
        j = 0
        colpaco = dfdict[tables[6]].columns.tolist()
        richmonth = dfo.richmonth(dfdict[tables[6]], tables[6], mdict, group[grcount].BACINO, colpaco, contract)
        for i in range(1, 4):
            updatekpi["FEH"].append(dict())
            updatekpi["FEH"][i]["name"] = target[i].name
            updatekpi["FEH"][i]["target"] = target[i].target
            updatekpi["FEH"][i]["tier1"] = target[i].tier1
            updatekpi["FEH"][i]["tier2"] = target[i].tier2
            updatekpi["FEH"][i]["tier3"] = target[i].tier3
            updatekpi["FEH"][i]["kpo"] = richmonth[j]["Rich_3_gg_p"]
            j += 1
        j = 0
        colopera = dfdict[tables[4]].columns.tolist()
        kpomonth = dfo.KpiFEHome2017(dfdict[tables[4]], colopera)
        for i in range(4, 12):
            updatekpi["FEH"].append(dict())
            updatekpi["FEH"][i]["name"] = target[i].name
            updatekpi["FEH"][i]["target"] = target[i].target
            updatekpi["FEH"][i]["tier1"] = target[i].tier1
            updatekpi["FEH"][i]["tier2"] = target[i].tier2
            updatekpi["FEH"][i]["tier3"] = target[i].tier3
            updatekpi["FEH"][i]["kpo"] = kpomonth[j]
            j += 1
        updatekpi["FEO"].append(dict())
        updatekpi["FEO"][0]["name"] = target[12].name
        updatekpi["FEO"][0]["target"] = target[12].target
        updatekpi["FEO"][0]["tier1"] = target[12].tier1
        updatekpi["FEO"][0]["tier2"] = target[12].tier2
        updatekpi["FEO"][0]["tier3"] = target[12].tier3
        updatekpi["FEO"][0]["kpo"] = ivrsmonth["ced_c87_ivr_191"][0][0]["Media_DD7"]
        j = 0
        colpaco = dfdict[tables[7]].columns.tolist()
        richmonth = dfo.richmonth(dfdict[tables[7]], tables[7], mdict, group[grcount].BACINO, colpaco, contract)
        for i in range(1, 2):
            updatekpi["FEO"].append(dict())
            updatekpi["FEO"][i]["name"] =  target[12+i].name
            updatekpi["FEO"][i]["target"] = target[12+i].target
            updatekpi["FEO"][i]["tier1"] = target[12+i].tier1
            updatekpi["FEO"][i]["tier2"] = target[12+i].tier2
            updatekpi["FEO"][i]["tier3"] = target[12+i].tier3
            updatekpi["FEO"][i]["kpo"] = richmonth[j]["Rich_3_gg_p"]
            j += 1
        j = 0
        colopera = dfdict[tables[5]].columns.tolist()
        kpomonth = dfo.KpiFEHome2017(dfdict[tables[5]], colopera)
        for i in range(2, 8):
            updatekpi["FEO"].append(dict())
            updatekpi["FEO"][i]["name"] =  target[12+i].name
            updatekpi["FEO"][i]["target"] = target[12+i].target
            updatekpi["FEO"][i]["tier1"] = target[12+i].tier1
            updatekpi["FEO"][i]["tier2"] = target[12+i].tier2
            updatekpi["FEO"][i]["tier3"] = target[12+i].tier3
            updatekpi["FEO"][i]["kpo"] = kpomonth[j]
            j += 1
        j = 0
        colopera = dfdict[tables[2]].columns.tolist()
        kpomonth = dfo.KpiBOHoOf2017(dfdict[tables[2]], dfdict[tables[3]], colopera)
        for i in range(0, 7):
            updatekpi["BO"].append(dict())
            updatekpi["BO"][i]["name"] = target[20+i].name
            updatekpi["BO"][i]["target"] = target[20+i].target
            updatekpi["BO"][i]["tier1"] = target[20+i].tier1
            updatekpi["BO"][i]["tier2"] = target[20+i].tier2
            updatekpi["BO"][i]["tier3"] = target[20+i].tier3
            updatekpi["BO"][i]["kpo"] = kpomonth[j]
            j += 1
        k = 0
        for i in typeupd:
            for j in range(0, len(updatekpi[i])):
                if updatekpi[i][j]["kpo"] != 0:
                    if target[j+k].kpi_type.id == 1:
                        if updatekpi[i][j]["target"] == 0:
                            updatekpi[i][j]["delta"] = updatekpi[i][j]["kpo"] - updatekpi[i][j]["tier1"]
                        else:
                            updatekpi[i][j]["delta"] = updatekpi[i][j]["kpo"] - updatekpi[i][j]["target"]
                    elif target[j+k].kpi_type.id == 2:
                        if updatekpi[i][j]["target"] == 0:
                            updatekpi[i][j]["delta"] = updatekpi[i][j]["tier1"] - updatekpi[i][j]["kpo"]
                        else:
                            updatekpi[i][j]["delta"] = updatekpi[i][j]["target"] - updatekpi[i][j]["kpo"]
                else:
                    updatekpi[i][j]["delta"] = 0
                updatekpi[i][j]["kpi_type"] = target[j+k].kpi_type.id
            k += j+1

        today = dt.datetime.now().date()
        x = 1

        try:
            if lastday[0].Day != today:
                dflastkpo = pd.read_sql("SELECT * FROM ced_c87_kpi_kpo_obtained WHERE MonthYear = '%s' AND substring(id,1,3) = '%s' AND Day = '%s'"
                                        % (date, group[grcount].ID, dt.datetime.strftime(lastday[0].Day, "%Y-%m-%d")), connection)
            else:
                dflastkpo = pd.read_sql("SELECT * FROM ced_c87_kpi_kpo_obtained WHERE MonthYear = '%s' AND "
                                        "substring(id,1,3) = '%s' AND Day < '%s' AND Day >= '%s'"
                                        % (date, group[grcount].ID, dt.datetime.strftime(today, "%Y-%m-%d"),
                                            dt.datetime.strftime(today - dt.timedelta(x), "%Y-%m-%d")),
                                        connection)
        except:
            dflastkpo = 0
        z = 3
        if dflastkpo != 0:
            collastkpo = dflastkpo.columns.tolist()
            for i in range(0, len(updatekpi)):
                for j in range(0, len(updatekpi[typeupd[i]])):
                    updatekpi[typeupd[i]][j]["old_kpo"] = dflastkpo[collastkpo[z]][0]
                    z += 1

        z = 0
        for i in range(0, len(updatekpi)):
            for j in range(0, len(updatekpi[typeupd[i]])):
                updatekpi[typeupd[i]][j]["malus_target"] = 0
                updatekpi[typeupd[i]][j]["malus_tier1"] = 0
                if target[z].kpi_type_id == 1:
                    if target[z].target != 0:
                        updatekpi[typeupd[i]][j]["malus_target"] = min(updatekpi[typeupd[i]][j]["kpo"] - target[z].target, 0)
                    if target[z].tier1 != 0:
                        updatekpi[typeupd[i]][j]["malus_tier1"] = min(updatekpi[typeupd[i]][j]["kpo"] - target[z].tier1 - updatekpi[typeupd[i]][j]["malus_target"], 0)
                    if updatekpi[typeupd[i]][j]["kpo"] < target[z].tier1:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier1
                    elif updatekpi[typeupd[i]][j]["kpo"] < target[z].tier2:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier2
                    elif updatekpi[typeupd[i]][j]["kpo"] < target[z].tier3:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier3
                    else:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier4

                elif target[z].kpi_type_id == 2:
                    if target[z].target != 0:
                        updatekpi[typeupd[i]][j]["malus_target"] = min(-(updatekpi[typeupd[i]][j]["kpo"] - target[z].target), 0)
                    if target[z].tier1 != 0:
                        updatekpi[typeupd[i]][j]["malus_tier1"] = min(target[z].tier1 - updatekpi[typeupd[i]][j]["kpo"] - updatekpi[typeupd[i]][j]["malus_target"], 0)
                    if updatekpi[typeupd[i]][j]["kpo"] > target[z].tier1:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier1
                    elif updatekpi[typeupd[i]][j]["kpo"] > target[z].tier2:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier2
                    elif updatekpi[typeupd[i]][j]["kpo"] > target[z].tier3:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier3
                    else:
                        updatekpi[typeupd[i]][j]["bonusmalus"] = target[z].bonustier4


                z += 1
    except Exception as error:
        print(type(error), error)
        updatekpi = {}
    return updatekpi
