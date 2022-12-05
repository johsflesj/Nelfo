import pandas as pd
import requests, openpyxl, math

bygningsmodell_excel = pd.read_excel("C:/Prosjekter/NELFO_solcelle/input_fag/utnyttingsgrad_og_bygningsmodeller.xlsx", 'bygningsmodeller')

bygningsmodell_smahus_lengde = bygningsmodell_excel['Lengde'][1]
bygningsmodell_boligblokk_lengde = bygningsmodell_excel['Lengde'][2]
bygningsmodell_barnehage_lengde = bygningsmodell_excel['Lengde'][3]
bygningsmodell_kontor_lengde = bygningsmodell_excel['Lengde'][4]
bygningsmodell_skole_lengde = bygningsmodell_excel['Lengde'][5]
bygningsmodell_universitet_lengde = bygningsmodell_excel['Lengde'][6]
bygningsmodell_sykehus_lengde = bygningsmodell_excel['Lengde'][7]
bygningsmodell_sykehjem_lengde = bygningsmodell_excel['Lengde'][8]
bygningsmodell_hotell_lengde = bygningsmodell_excel['Lengde'][9]
bygningsmodell_idrettsbygning_lengde = bygningsmodell_excel['Lengde'][10]
bygningsmodell_forretning_lengde = bygningsmodell_excel['Lengde'][11]
bygningsmodell_kultur_lengde = bygningsmodell_excel['Lengde'][12]
bygningsmodell_industri_lengde = bygningsmodell_excel['Lengde'][13]


bygningsmodell_smahus_bredde = bygningsmodell_excel['Bredde'][1]
bygningsmodell_boligblokk_bredde = bygningsmodell_excel['Bredde'][2]
bygningsmodell_barnehage_bredde = bygningsmodell_excel['Bredde'][3]
bygningsmodell_kontor_bredde = bygningsmodell_excel['Bredde'][4]
bygningsmodell_skole_bredde = bygningsmodell_excel['Bredde'][5]
bygningsmodell_universitet_bredde = bygningsmodell_excel['Bredde'][6]
bygningsmodell_sykehus_bredde = bygningsmodell_excel['Bredde'][7]
bygningsmodell_sykehjem_bredde = bygningsmodell_excel['Bredde'][8]
bygningsmodell_hotell_bredde = bygningsmodell_excel['Bredde'][9]
bygningsmodell_idrettsbygning_bredde = bygningsmodell_excel['Bredde'][10]
bygningsmodell_forretning_bredde = bygningsmodell_excel['Bredde'][11]
bygningsmodell_kultur_bredde = bygningsmodell_excel['Bredde'][12]
bygningsmodell_industri_bredde = bygningsmodell_excel['Bredde'][13]


bygningsmodell_smahus_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][1]
bygningsmodell_boligblokk_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][2]
bygningsmodell_barnehage_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][3]
bygningsmodell_kontor_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][4]
bygningsmodell_skole_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][5]
bygningsmodell_universitet_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][6]
bygningsmodell_sykehus_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][7]
bygningsmodell_sykehjem_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][8]
bygningsmodell_hotell_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][9]
bygningsmodell_idrettsbygning_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][10]
bygningsmodell_forretning_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][11]
bygningsmodell_kultur_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][12]
bygningsmodell_industri_etasjehoyde = bygningsmodell_excel['Etasjehoyde'][13]

bygningsmodell_smahus_VinduDor = bygningsmodell_excel['Andel vindu og dor'][1]
bygningsmodell_boligblokk_VinduDor = bygningsmodell_excel['Lengde'][2]
bygningsmodell_barnehage_VinduDor = bygningsmodell_excel['Lengde'][3]
bygningsmodell_kontor_VinduDor = bygningsmodell_excel['Lengde'][4]
bygningsmodell_skole_VinduDor = bygningsmodell_excel['Lengde'][5]
bygningsmodell_universitet_VinduDor = bygningsmodell_excel['Lengde'][6]
bygningsmodell_sykehus_VinduDor = bygningsmodell_excel['Lengde'][7]
bygningsmodell_sykehjem_VinduDor = bygningsmodell_excel['Lengde'][8]
bygningsmodell_hotell_VinduDor = bygningsmodell_excel['Lengde'][9]
bygningsmodell_idrettsbygning_VinduDor = bygningsmodell_excel['Lengde'][10]
bygningsmodell_forretning_VinduDor = bygningsmodell_excel['Lengde'][11]
bygningsmodell_kultur_VinduDor = bygningsmodell_excel['Lengde'][12]
bygningsmodell_industri_VinduDor = bygningsmodell_excel['Lengde'][13]

print(math.cos(6))