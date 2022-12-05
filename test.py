import pandas as pd
import requests, openpyxl


input_utnytting_bygningsmodell = pd.read_excel("C:/Prosjekter/NELFO_solcelle/input_fag/utnyttingsgrad_og_bygningsmodeller.xlsx")

#SKRÅTAK
utnyttingsgrad_skratak_enebolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][0] ,3)
utnyttingsgrad_skratak_tomannsbolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][1] ,3)
utnyttingsgrad_skratak_rekkehus = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][2] ,3)
utnyttingsgrad_skratak_storeboliger = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][3] ,3)
utnyttingsgrad_skratak_bofelleskap = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][4] ,3)
utnyttingsgrad_skratak_fritidsbolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][5] ,3)
utnyttingsgrad_skratak_koie = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][6] ,3)
utnyttingsgrad_skratak_garasje = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][7] ,3)
utnyttingsgrad_skratak_annenbolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][8] ,3)
utnyttingsgrad_skratak_industri = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][9] ,3)
utnyttingsgrad_skratak_lager = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][10] ,3)
utnyttingsgrad_skratak_fiskeri = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][11] ,3)
utnyttingsgrad_skratak_kontor = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][12] ,3)
utnyttingsgrad_skratak_forretning = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][13] ,3)
utnyttingsgrad_skratak_messe = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][14] ,3)
utnyttingsgrad_skratak_ekspedisjonterminal = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][15] ,3)
utnyttingsgrad_skratak_garasjehangar = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][16] ,3)
utnyttingsgrad_skratak_vegtrafikk = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][17] ,3)
utnyttingsgrad_skratak_hotell = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][18] ,3)
utnyttingsgrad_skratak_overnatting = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][19] ,3)
utnyttingsgrad_skratak_restaurant = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][20] ,3)
utnyttingsgrad_skratak_skole = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][21] ,3)
utnyttingsgrad_skratak_universitet = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][22] ,3)
utnyttingsgrad_skratak_museum = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][23] ,3)
utnyttingsgrad_skratak_idrett = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][24] ,3)
utnyttingsgrad_skratak_kultur = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][25] ,3)
utnyttingsgrad_skratak_religios = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][26] ,3)
utnyttingsgrad_skratak_sykehus = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][27] ,3)
utnyttingsgrad_skratak_sykehjem = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][28] ,3)
utnyttingsgrad_skratak_primarhelse = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][29] ,3)
utnyttingsgrad_skratak_beredskap = round(input_utnytting_bygningsmodell['utnyttingsgrad_skratak'][30] ,3)


#FLATT TAK
utnyttingsgrad_flatt_tak_enebolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][0] ,3)
utnyttingsgrad_flatt_tak_tomannsbolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][1] ,3)
utnyttingsgrad_flatt_tak_rekkehus = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][2] ,3)
utnyttingsgrad_flatt_tak_storeboliger = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][3] ,3)
utnyttingsgrad_flatt_tak_bofelleskap = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][4] ,3)
utnyttingsgrad_flatt_tak_fritidsbolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][5] ,3)
utnyttingsgrad_flatt_tak_koie = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][6] ,3)
utnyttingsgrad_flatt_tak_garasje = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][7] ,3)
utnyttingsgrad_flatt_tak_annenbolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][8] ,3)
utnyttingsgrad_flatt_tak_industri = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][9] ,3)
utnyttingsgrad_flatt_tak_lager = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][10] ,3)
utnyttingsgrad_flatt_tak_fiskeri = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][11] ,3)
utnyttingsgrad_flatt_tak_kontor = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][12] ,3)
utnyttingsgrad_flatt_tak_forretning = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][13] ,3)
utnyttingsgrad_flatt_tak_messe = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][14] ,3)
utnyttingsgrad_flatt_tak_ekspedisjonterminal = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][15] ,3)
utnyttingsgrad_flatt_tak_garasjehangar = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][16] ,3)
utnyttingsgrad_flatt_tak_vegtrafikk = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][17] ,3)
utnyttingsgrad_flatt_tak_hotell = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][18] ,3)
utnyttingsgrad_flatt_tak_overnatting = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][19] ,3)
utnyttingsgrad_flatt_tak_restaurant = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][20] ,3)
utnyttingsgrad_flatt_tak_skole = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][21] ,3)
utnyttingsgrad_flatt_tak_universitet = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][22] ,3)
utnyttingsgrad_flatt_tak_museum = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][23] ,3)
utnyttingsgrad_flatt_tak_idrett = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][24] ,3)
utnyttingsgrad_flatt_tak_kultur = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][25] ,3)
utnyttingsgrad_flatt_tak_religios = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][26] ,3)
utnyttingsgrad_flatt_tak_sykehus = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][27] ,3)
utnyttingsgrad_flatt_tak_sykehjem = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][28] ,3)
utnyttingsgrad_flatt_tak_primarhelse = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][29] ,3)
utnyttingsgrad_flatt_tak_beredskap = round(input_utnytting_bygningsmodell['utnyttingsgrad_flatt_tak'][30] ,3)


#VEGG
utnyttingsgrad_vegg_enebolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][0] ,3)
utnyttingsgrad_vegg_tomannsbolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][1] ,3)
utnyttingsgrad_vegg_rekkehus = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][2] ,3)
utnyttingsgrad_vegg_storeboliger = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][3] ,3)
utnyttingsgrad_vegg_bofelleskap = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][4] ,3)
utnyttingsgrad_vegg_fritidsbolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][5] ,3)
utnyttingsgrad_vegg_koie = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][6] ,3)
utnyttingsgrad_vegg_garasje = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][7] ,3)
utnyttingsgrad_vegg_annenbolig = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][8] ,3)
utnyttingsgrad_vegg_industri = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][9] ,3)
utnyttingsgrad_vegg_lager = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][10] ,3)
utnyttingsgrad_vegg_fiskeri = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][11] ,3)
utnyttingsgrad_vegg_kontor = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][12] ,3)
utnyttingsgrad_vegg_forretning = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][13] ,3)
utnyttingsgrad_vegg_messe = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][14] ,3)
utnyttingsgrad_vegg_ekspedisjonterminal = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][15] ,3)
utnyttingsgrad_vegg_garasjehangar = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][16] ,3)
utnyttingsgrad_vegg_vegtrafikk = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][17] ,3)
utnyttingsgrad_vegg_hotell = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][18] ,3)
utnyttingsgrad_vegg_overnatting = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][19] ,3)
utnyttingsgrad_vegg_restaurant = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][20] ,3)
utnyttingsgrad_vegg_skole = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][21] ,3)
utnyttingsgrad_vegg_universitet = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][22] ,3)
utnyttingsgrad_vegg_museum = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][23] ,3)
utnyttingsgrad_vegg_idrett = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][24] ,3)
utnyttingsgrad_vegg_kultur = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][25] ,3)
utnyttingsgrad_vegg_religios = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][26] ,3)
utnyttingsgrad_vegg_sykehus = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][27] ,3)
utnyttingsgrad_vegg_sykehjem = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][28] ,3)
utnyttingsgrad_vegg_primarhelse = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][29] ,3)
utnyttingsgrad_vegg_beredskap = round(input_utnytting_bygningsmodell['utnyttingsgrad_vegg'][30] ,3)

takvinkel_enebolig = input_utnytting_bygningsmodell['Takvinkel'][0]
takvinkel_tomannsbolig = input_utnytting_bygningsmodell['Takvinkel'][1]
takvinkel_rekkehus = input_utnytting_bygningsmodell['Takvinkel'][2]
takvinkel_storeboliger = input_utnytting_bygningsmodell['Takvinkel'][3]
takvinkel_bofelleskap = input_utnytting_bygningsmodell['Takvinkel'][4]
takvinkel_fritidsbolig = input_utnytting_bygningsmodell['Takvinkel'][5]
takvinkel_koie = input_utnytting_bygningsmodell['Takvinkel'][6]
takvinkel_garasje = input_utnytting_bygningsmodell['Takvinkel'][7]
takvinkel_annenbolig = input_utnytting_bygningsmodell['Takvinkel'][8]
takvinkel_industri = input_utnytting_bygningsmodell['Takvinkel'][9]
takvinkel_lager = input_utnytting_bygningsmodell['Takvinkel'][10]
takvinkel_fiskeri = input_utnytting_bygningsmodell['Takvinkel'][11]
takvinkel_kontor = input_utnytting_bygningsmodell['Takvinkel'][12]
takvinkel_forretning = input_utnytting_bygningsmodell['Takvinkel'][13]
takvinkel_messe = input_utnytting_bygningsmodell['Takvinkel'][14]
takvinkel_ekspedisjonterminal = input_utnytting_bygningsmodell['Takvinkel'][15]
takvinkel_garasjehangar = input_utnytting_bygningsmodell['Takvinkel'][16]
takvinkel_vegtrafikk = input_utnytting_bygningsmodell['Takvinkel'][17]
takvinkel_hotell = input_utnytting_bygningsmodell['Takvinkel'][18]
takvinkel_overnatting = input_utnytting_bygningsmodell['Takvinkel'][19]
takvinkel_restaurant = input_utnytting_bygningsmodell['Takvinkel'][20]
takvinkel_skole = input_utnytting_bygningsmodell['Takvinkel'][21]
takvinkel_universitet = input_utnytting_bygningsmodell['Takvinkel'][22]
takvinkel_museum = input_utnytting_bygningsmodell['Takvinkel'][23]
takvinkel_idrett = input_utnytting_bygningsmodell['Takvinkel'][24]
takvinkel_kultur = input_utnytting_bygningsmodell['Takvinkel'][25]
takvinkel_religios = input_utnytting_bygningsmodell['Takvinkel'][26]
takvinkel_sykehus = input_utnytting_bygningsmodell['Takvinkel'][27]
takvinkel_sykehjem = input_utnytting_bygningsmodell['Takvinkel'][28]
takvinkel_primarhelse = input_utnytting_bygningsmodell['Takvinkel'][29]
takvinkel_beredskap = input_utnytting_bygningsmodell['Takvinkel'][30]




