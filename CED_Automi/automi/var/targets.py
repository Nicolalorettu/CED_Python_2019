# c87 targets
oldkpic87 = {"FE_RicS": 24.49, "FE_RicMC": 22.99, "FE_RicC": 24.99,
             "FE_ORic": 28.49,
             "FE_RHF": 14.99, "FE_RHD": 16.49, "FE_IBOHF": 11.99,
             "FE_IBOHD": 14.99, "FE_IOFHF": 26.99, "FE_IOFHD": 13.49,
             "FE_RO": 14.49, "FE_IBOO": 15.49, "FE_IOFO": 21.99,
             "BO_RF": 21.99, "BO_RA": 21.99, "BO_RF": 21.99,
             "BO_IOFHF": 27.49, "BO_IOFHA": 12.99, "BO_IOFHFi": 11.99,
             "BO_CTFD": 15.99}

old_opera_kpi_names = {"FEH": ["Rework Home Fonia", "Rework Home Dati",
                               "Invio BO Home Fonia", "Invio BO Home Dati",
                               "Invio OF Home Fonia", "Invio OF Home Dati"],
                       "FEO": ["Rework Office", "Invio BO Office", "Invio OF Office"],
                       "BO": ["Rework Fonia", "Rework ADSL", "Rework Fibra",
                              "Invio OF Home Fonia", "Invio OF Home ADSL",
                              "Invio OF Home FIBRA", "C TEAM FONIA + DATI"]}

old_opera_kpi_targets = {"FEH": [14.99, 16.49, 11.99, 14.99, 26.99, 13.49],
                         "FEO": [14.49, 15.49, 21.99],
                         "BO": [21.99, 21.99, 21.99, 27.49, 12.99, 11.99, 15.99]}

new_opera_kpi_names = {"FEH": ["Rework Home Fonia", "Rework Home Dati",
                               "Rework Home Fibra", "ONT Fonia", "ONT ADSL",
                               "ONT Fibra", "Invio BO Home Fonia",
                               "Invio BO Home Dati", "Invio BO Home Fibra",
                               "Invio OF Home Fonia", "Invio OF Home Dati",
                               "Invio OF Home Fibra"],
                       "FEO": ["Rework Medie", "Rework Complesse", "ONT Medie",
                               "ONT Complesse", "Invio BO Medie",
                               "Invio BO Complesse", "Invio OF Medie",
                               "Invio OF Complesse"],
                       "BO": ["BO REWORK (VN+RIP7GG)FONIA", "BO REWORK (VN+RIP7GG)ADSL",
                              "BO REWORK (VN+RIP7GG)FIBRA", "ONT Fonia", "ONT ADSL",
                              "ONT Fibra", "BO Invio OF Fonia", "BO Invio OF ADSL",
                              "BO Invio OF Fibra", "C TEAM FONIA + DATI",
                              "RIPETIZIONE A 33 GG SU COLLAUDI"]}

new_opera_kpi_targets = {"FEH": [15.5, 20, 15.5, 7, 8.5, 6.5, 14, 14, 15, 28, 16, 11],
                         "FEO": [16.5, 15.5, 9, 8, 16, 15, 22, 11],
                         "BO": [15.5, 15.5, 15.5, 8.5, 9, 6.5, 18, 14, 10, 0, 0]}

newtargetrichHOME = [22.99, 22.99, 22.99]
newtargetrichOFF = [22.99, 21.99]

newkpic87 = {"FE_RicS": 22.99, "FE_RicMC": 22.99, "FE_RicC": 22.99,
             "FE_ORicMC": 22.99, "FE_ORicC": 21.99,
             "FE_RHF": 15.5, "FE_RHD": 20, "FE_RHFi": 15.5,
             "FE_ONT_FH": 7, "FE_ONT_DH": 8.5, "FE_ONT_FiH": 6.5,
             "FE_IBOHF": 14, "FE_IBOHD": 14, "FE_IBOHFi": 15,
             "FE_IOFHF": 28, "FE_IOFHD": 16, "FE_IOFHFi": 11,
             "FE_RMOF": 16.5, "FE_RCOF": 15.5, "FE_ONT_MOF": 9,
             "FE_ONT_COF": 8, "FE_IBOOFM": 16, "FE_IBOOFC": 15,
             "FE_IOFOFM": 22, "FE_IOFOFC": 11, "BO_VR7F": 15.5,
             "BO_VR7D": 15.5, "BO_VR7Fi": 15.5, "BO_ONT_F": 8.5,
             "BO_ONT_D": 9, "BO_ONT_Fi": 6.5, "BO_IOFF": 18,
             "BO_IOFD": 14, "BO_IOFFi": 10, "BOCTFD": 16,
             "BO_Rip_3GCol": 4.8}

newivrc87 = {"Semplici": [6.51, 7.51, 8.51], "MedioCom": [7.51, 8.51, 9.01], "Complesse": [8.01, 9.01, 9.51]}

oldivrc87 = {"Semplici": 7.41, "MedioCom": 7.41, "Complesse": 7.41}

# TABLES var

columnkpo = { "FEH": ["Ivr Semplici", "Ivr Medio Comp.", "Ivr Complesse", "Richiamate Semplici",
                     "Richiamate Medio Comp.", "Richiamate Complesse", "Rework FONIA HOME",
                     "Rework DATI HOME", "Rework FIBRA HOME", "ONT Fonia", "ONT ADSL", "ONT Fibra",
                     "Invio BO Fonia", "Invio BO ADSL", "Invio BO Fibra", "Invio OF Fonia",
                     "Invio OF ADSL", "Invio OF Fibra"],
             "FEO": ["Ivr Medio Comp.", "Ivr Complesse", "Richiamate Medio Comp.",
                    "Richiamate Complesse", "Rework Medie", "Rework Complesse", "ONT Medie",
                    "ONT Complesse", "Invio BO Medie", "Invio BO Complesse", "Invio OF Medie",
                    "Invio OF Complesse"],
             "BO": ["BO REWORK (VN+RIP3GG)FONIA", "BO REWORK (VN+RIP3GG)ADSL",
                   "BO REWORK (VN+RIP3GG)FIBRA", "ONT Fonia", "ONT ADSL", "ONT Fibra",
                   "BO INVIO OF FONIA", "BO INVIO OF ADSL", "BO INVIO OF FIBRA",
                   "C TEAM FONIA + DATI", "RIPETIZIONE A 33 GG SU COLLAUDI"]}


operadb = {"MonthYear": [["1", "4"]]}
ivrdb = {"DATA_INT": [["4", "2"], ["9", "2"]]}
pacodb = {"Data": [["4", "2"], ["7", "2"]]}
