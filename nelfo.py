#Programmet leser excel-fil som er eksport av GIS-data. Denne filen heter:
#Programmet henter eksisterende excel-mal og skriver til ny fil.

# OBS! Programmet må oppdateres med riktig lenke til filer som programmet leser og skriver til.

#Biblioteker som må installeres: Pandas, requests, xlsxwriter, openpyxl


#importere nødvendige bibliotek
import pandas as pd
import requests, openpyxl, math

#Excel-fil som inneholder GIS-uttrekk
kildedata = pd.read_excel("C:/Prosjekter/NELFO_solcelle/fme_excel_export.xlsx")

#Setter alle variabler
bebygdareal = kildedata['bebygdareal']
fkbareal = kildedata['fkbareal']
bruksarealtilbolig = kildedata['bruksarealtilbolig']
bruksarealtilannet = kildedata['bruksarealtilannet']
alternativtareal = kildedata['alternativtareal']
alternativtareal2 = kildedata['alternativtareal2']
utenbebygdareal = kildedata['utenbebygdareal']
antalletasjer = kildedata['antalletasjer']
elspotomr = kildedata['ElSpotOmr']
bygningstypekode_tekst_tre = kildedata['bygningstypekode_tekst_tre']
tellekolonne = kildedata['Tellekolonne']
bygningskode_tekst_to = kildedata['bygningskode_tekst_to']
name = kildedata['name']
shape_length = kildedata['Shape_Length']
shape_area = kildedata['Shape_Area']
kommunenummer = kildedata['kommunenummer']
kommunenavn = kildedata['navn']
fylkesnummer = kildedata['fylkesnummer']
fid = kildedata['fid']

#Excel-fil vi skal skrive data til:
regneark_eksisterende = "C:/Prosjekter/NELFO_solcelle/tom_test.xlsx"
wb = openpyxl.load_workbook(regneark_eksisterende)

antall = len(kildedata) #Antall rader i exceldokument.




########################################################################
# Utnyttingsgrad og bygningsmodell

# Ecel.fil som innegolder utnyttingsgrad og bygningsmodell.
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

##### TAKVINKEL
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

##########################################################################################
##### INPUT FRA BYGNINGSMODELL
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


bygningsmodell_smahus_bredde = bygningsmodell_excel['Lengde'][1]
bygningsmodell_boligblokk_bredde = bygningsmodell_excel['Lengde'][2]
bygningsmodell_barnehage_bredde = bygningsmodell_excel['Lengde'][3]
bygningsmodell_kontor_bredde = bygningsmodell_excel['Lengde'][4]
bygningsmodell_skole_bredde = bygningsmodell_excel['Lengde'][5]
bygningsmodell_universitet_bredde = bygningsmodell_excel['Lengde'][6]
bygningsmodell_sykehus_bredde = bygningsmodell_excel['Lengde'][7]
bygningsmodell_sykehjem_bredde = bygningsmodell_excel['Lengde'][8]
bygningsmodell_hotell_bredde = bygningsmodell_excel['Lengde'][9]
bygningsmodell_idrettsbygning_bredde = bygningsmodell_excel['Lengde'][10]
bygningsmodell_forretning_bredde = bygningsmodell_excel['Lengde'][11]
bygningsmodell_kultur_bredde = bygningsmodell_excel['Lengde'][12]
bygningsmodell_industri_bredde = bygningsmodell_excel['Lengde'][13]


bygningsmodell_smahus_etasjehoyde = bygningsmodell_excel['Lengde'][1]
bygningsmodell_boligblokk_etasjehoyde = bygningsmodell_excel['Lengde'][2]
bygningsmodell_barnehage_etasjehoyde = bygningsmodell_excel['Lengde'][3]
bygningsmodell_kontor_etasjehoyde = bygningsmodell_excel['Lengde'][4]
bygningsmodell_skole_etasjehoyde = bygningsmodell_excel['Lengde'][5]
bygningsmodell_universitet_etasjehoyde = bygningsmodell_excel['Lengde'][6]
bygningsmodell_sykehus_etasjehoyde = bygningsmodell_excel['Lengde'][7]
bygningsmodell_sykehjem_etasjehoyde = bygningsmodell_excel['Lengde'][8]
bygningsmodell_hotell_etasjehoyde = bygningsmodell_excel['Lengde'][9]
bygningsmodell_idrettsbygning_etasjehoyde = bygningsmodell_excel['Lengde'][10]
bygningsmodell_forretning_etasjehoyde = bygningsmodell_excel['Lengde'][11]
bygningsmodell_kultur_etasjehoyde = bygningsmodell_excel['Lengde'][12]
bygningsmodell_industri_etasjehoyde = bygningsmodell_excel['Lengde'][13]

bygningsmodell_smahus_VinduDor = bygningsmodell_excel['Lengde'][1]
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


#Henter alle kommunenummer fra kommunenummerregister i Geonorge (Kartverket) som kan itereres over. Kartverket sørger for at kodelisten er oppdatert i henhold til gjeldende inndelinger
kom_url = "https://register.geonorge.no/api/sosi-kodelister/kommunenummer.json?" #URL til kodeliste
r = requests.get(kom_url) #Kobler opp til url
kom_data = r.json() #henter innhold som JSON

kom_antall = len(kom_data['containeditems']) #henter alle kommuner i kodelisten

#Henter alle fylkesnummer fra register i Geonorge (Kartverket) som kan itereres over. Kartverket sørger for at kodelisten er oppdatert i henhold til gjeldende inndelinger.
fylke_url = "https://register.geonorge.no/api/sosi-kodelister/fylkesnummer.json?"
r_fylke = requests.get(fylke_url)
fylke_data = r_fylke.json()
fylke_antall = len(fylke_data['containeditems'])

fylke = [] #Liste som populeres med alle fylkesnummer utenom Svalbard og Jan Mayen

for fylkekode in range(fylke_antall): #Loop som looper gjennom fylke_url og henter alle fylkesnummer
    fylkenr = fylke_data['containeditems'][fylkekode]['codevalue'] #Lokasjon for fylkesnummer i json-fil (fylke_url)
    if fylkenr != "21" and fylkenr != "22": #Utelukker Svalbard og Jan Mayen
        fylke.append(fylkenr)

fylke.sort() #Sorterer fylkeslisten slik at minste fylkesnummer kommer først.

fylke = ["50"]# !!! OBS KUN TIL TEST. FJERNES I PRODUKSJON!

###KWH!!
for f in range(1): #!!!!! MÅ BYTTES TIL len(fylke)
    row = 0
    for kom in range(kom_antall): #Looper gjennom kodelisten for hver kommune
        kom_nr = kom_data['containeditems'][kom]['codevalue'] #kommunenummer
        kom_navn = kom_data['containeditems'][kom]['label']
        if str(fylke[f]) == str(kom_nr[:2]):
            row += 1
            print("fylke", fylke[f])

            #Tellere for hver attributt
            bebygdareal_tot = 0
            fkbareal_tot = 0
            bebygdareal_tot = 0 
            fkbareal_tot = 0 
            bruksarealtilbolig_tot = 0 
            bruksarealtilannet_tot = 0 
            alternativtareal_tot = 0 
            alternativtareal2_tot = 0 
            utenbebygdareal_tot = 0 
            antalletasjer_tot = 0
            bygningstypekode_tekst_tre_tot = 0 
            tellekolonne_tot = 0 
            bygningskode_tekst_to_tot = 0
            shape_length_tot = 0 
            shape_area_tot = 0 

            #Bygningstyper: Oppsumerte variabler for areal av hver bygningstype
            btype_enebolig_areal = 0
            btype_tomannsbolig_areal = 0
            btype_rekkehus_areal = 0
            btype_storeboliger_areal = 0
            btype_bofellesskap_areal = 0
            btype_fritidsbolig_areal = 0
            btype_koie_areal = 0
            btype_garasje_areal = 0
            btype_annenbolig_areal = 0
            btype_industri_areal = 0
            btype_lager_areal = 0
            btype_fiskeri_areal = 0
            btype_kontor_areal = 0
            btype_forretning_areal = 0
            btype_messekongress_areal = 0
            btype_terminal_areal = 0
            btype_garasjehangar_areal = 0
            btype_vegtrafikk_areal = 0
            btype_hotell_areal = 0
            btype_overnatting_areal = 0
            btype_restaurant_areal = 0
            btype_skole_areal = 0
            btype_universitet_areal = 0
            btype_museum_areal = 0
            btype_idrett_areal = 0
            btype_kulturhus_areal = 0
            btype_religios_areal = 0
            btype_sykehus_areal = 0
            btype_sykehjem_areal = 0
            btype_primarhelse_areal = 0
            btype_beredskap_areal = 0


            #Bygningstype: Oppsummerte variabler for antall av hver bygningstype
            btype_enebolig_antall = 0
            btype_tomannsbolig_antall = 0
            btype_rekkehus_antall = 0
            btype_storeboliger_antall = 0
            btype_bofellesskap_antall = 0
            btype_fritidsbolig_antall = 0
            btype_koie_antall = 0
            btype_garasje_antall = 0
            btype_annenbolig_antall = 0
            btype_industri_antall = 0
            btype_lager_antall = 0
            btype_fiskeri_antall = 0
            btype_kontor_antall = 0
            btype_forretning_antall = 0
            btype_messekongress_antall = 0
            btype_terminal_antall = 0
            btype_garasjehangar_antall = 0
            btype_vegtrafikk_antall = 0
            btype_hotell_antall = 0
            btype_overnatting_antall = 0
            btype_restaurant_antall = 0
            btype_skole_antall = 0
            btype_universitet_antall = 0
            btype_museum_antall = 0
            btype_idrett_antall = 0
            btype_kulturhus_antall = 0
            btype_religios_antall = 0
            btype_sykehus_antall = 0
            btype_sykehjem_antall = 0
            btype_primarhelse_antall = 0
            btype_beredskap_antall = 0

            
            #Oppsummerte variabler for sum etasjer per bygningstype
            btype_enebolig_etasje = 0
            btype_tomannsbolig_etasje = 0
            btype_rekkehus_etasje = 0
            btype_storeboliger_etasje = 0
            btype_bofellesskap_etasje = 0
            btype_fritidsbolig_etasje = 0
            btype_koie_etasje = 0
            btype_garasje_etasje = 0
            btype_annenbolig_etasje = 0
            btype_industri_etasje = 0
            btype_lager_etasje = 0
            btype_fiskeri_etasje = 0
            btype_kontor_etasje = 0
            btype_forretning_etasje = 0
            btype_messekongress_etasje = 0
            btype_terminal_etasje = 0
            btype_garasjehangar_etasje = 0
            btype_vegtrafikk_etasje = 0
            btype_hotell_etasje = 0
            btype_overnatting_etasje = 0
            btype_restaurant_etasje = 0
            btype_skole_etasje = 0
            btype_universitet_etasje = 0
            btype_museum_etasje = 0
            btype_idrett_etasje = 0
            btype_kulturhus_etasje = 0
            btype_religios_etasje = 0
            btype_sykehus_etasje = 0
            btype_sykehjem_etasje = 0
            btype_primarhelse_etasje = 0
            btype_beredskap_etasje = 0

            #Variabler som oppsummerer omkrets på bygningstype
            btype_enebolig_omkrets = 0
            btype_tomannsbolig_omkrets = 0
            btype_rekkehus_omkrets = 0
            btype_storeboliger_omkrets = 0
            btype_bofellesskap_omkrets = 0
            btype_fritidsbolig_omkrets = 0
            btype_koie_omkrets = 0
            btype_garasje_omkrets = 0
            btype_annenbolig_omkrets = 0
            btype_industri_omkrets = 0
            btype_lager_omkrets = 0
            btype_fiskeri_omkrets = 0
            btype_kontor_omkrets = 0
            btype_forretning_omkrets = 0
            btype_messekongress_omkrets = 0
            btype_terminal_omkrets = 0
            btype_garasjehangar_omkrets = 0
            btype_vegtrafikk_omkrets = 0
            btype_hotell_omkrets = 0
            btype_overnatting_omkrets = 0
            btype_restaurant_omkrets = 0
            btype_skole_omkrets = 0
            btype_universitet_omkrets = 0
            btype_museum_omkrets = 0
            btype_idrett_omkrets = 0
            btype_kulturhus_omkrets = 0
            btype_religios_omkrets = 0
            btype_sykehus_omkrets = 0
            btype_sykehjem_omkrets = 0
            btype_primarhelse_omkrets = 0
            btype_beredskap_omkrets = 0
                
            for i in range(antall):
                if str(kom_nr) == str(kommunenummer[i]):
                    
                    #print(bygningstypekode_tekst_tre[i])
            #         bebygdareal_tot += bebygdareal[i]
            #         fkbareal_tot += fkbareal[i]
            #         bruksarealtilbolig_tot += bruksarealtilbolig[i]
            #         bruksarealtilannet_tot += bruksarealtilannet[i]
            #         alternativtareal_tot += alternativtareal[i]
            #         alternativtareal2_tot += alternativtareal2[i]
            #         utenbebygdareal_tot += utenbebygdareal[i]
            #         antalletasjer_tot += antalletasjer[i]
            #         tellekolonne_tot += tellekolonne[i]
            #         shape_length_tot += shape_length[i]
            #         shape_area_tot += shape_area[i]

                    if bygningskode_tekst_to[i] == 11:
                        btype_enebolig_areal += bebygdareal[i]
                        btype_enebolig_antall += 1
                    elif bygningskode_tekst_to[i] == 12:
                        btype_tomannsbolig_areal += bebygdareal[i]
                        btype_tomannsbolig_antall += 1
                    elif bygningskode_tekst_to[i] == 13:
                        btype_rekkehus_areal += bebygdareal[i]
                        btype_rekkehus_antall += 1
                    elif bygningskode_tekst_to[i] == 14:
                        btype_storeboliger_areal += bebygdareal[i]
                        btype_storeboliger_antall += 1
                    elif bygningskode_tekst_to[i] == 15:
                        btype_bofellesskap_areal += bebygdareal[i]
                        btype_bofellesskap_antall += 1
                    elif bygningskode_tekst_to[i] == 16:
                        btype_fritidsbolig_areal += bebygdareal[i]
                        btype_fritidsbolig_antall += 1
                    elif bygningskode_tekst_to[i] == 17:
                        btype_koie_areal += bebygdareal[i]
                        btype_koie_antall += 1
                    elif bygningskode_tekst_to[i] == 18:
                        btype_garasje_areal += bebygdareal[i]
                        btype_garasje_antall += 1
                    elif bygningskode_tekst_to[i] == 19:
                        btype_annenbolig_areal += bebygdareal[i]
                        btype_annenbolig_antall += 1
                    elif bygningskode_tekst_to[i] == 21:
                        btype_industri_areal += bebygdareal[i]
                        btype_industri_antall += 1
                    elif bygningskode_tekst_to[i] == 23:
                        btype_lager_areal += bebygdareal[i]
                        btype_lager_antall += 1
                    elif bygningskode_tekst_to[i] == 24:
                        btype_fiskeri_areal += bebygdareal[i]
                        btype_fiskeri_antall += 1
                    elif bygningskode_tekst_to[i] == 31:
                        btype_kontor_areal += bebygdareal[i]
                        btype_kontor_antall += 1
                    elif bygningskode_tekst_to[i] == 32:
                        btype_forretning_areal += bebygdareal[i]
                        btype_forretning_antall += 1
                    elif bygningskode_tekst_to[i] == 33:
                        btype_messekongress_areal += bebygdareal[i]
                        btype_messekongress_antall += 1
                    elif bygningskode_tekst_to[i] == 41:
                        btype_terminal_areal += bebygdareal[i]
                        btype_terminal_antall += 1
                    elif bygningskode_tekst_to[i] == 43:
                        btype_garasjehangar_areal += bebygdareal[i]
                        btype_garasjehangar_antall += 1
                    elif bygningskode_tekst_to[i] == 44:
                        btype_vegtrafikk_areal += bebygdareal[i]
                        btype_vegtrafikk_antall += 1
                    elif bygningskode_tekst_to[i] == 51:
                        btype_hotell_areal += bebygdareal[i]
                        btype_hotell_antall += 1
                    elif bygningskode_tekst_to[i] == 52:
                        btype_overnatting_areal += bebygdareal[i]
                        btype_overnatting_antall += 1
                    elif bygningskode_tekst_to[i] == 53:
                        btype_restaurant_areal += bebygdareal[i]
                        btype_restaurant_antall += 1
                    elif bygningskode_tekst_to[i] == 61:
                        btype_skole_areal += bebygdareal[i]
                        btype_skole_antall += 1
                    elif bygningskode_tekst_to[i] == 62:
                        btype_universitet_areal += bebygdareal[i]
                        btype_universitet_antall += 1
                    elif bygningskode_tekst_to[i] == 64:
                        btype_museum_areal += bebygdareal[i]
                        btype_museum_antall += 1
                    elif bygningskode_tekst_to[i] == 65:
                        btype_idrett_areal += bebygdareal[i]
                        btype_idrett_antall += 1
                    elif bygningskode_tekst_to[i] == 66:
                        btype_kulturhus_areal += bebygdareal[i]
                        btype_kulturhus_antall += 1
                    elif bygningskode_tekst_to[i] == 67:
                        btype_religios_areal += bebygdareal[i]
                        btype_religios_antall += 1
                    elif bygningskode_tekst_to[i] == 71:
                        btype_sykehus_areal += bebygdareal[i]
                        btype_sykehus_antall += 1
                    elif bygningskode_tekst_to[i] == 72:
                        btype_sykehjem_areal += bebygdareal[i]
                        btype_sykehjem_antall += 1
                    elif bygningskode_tekst_to[i] == 73:
                        btype_primarhelse_areal += bebygdareal[i]
                        btype_primarhelse_antall += 1
                    elif bygningskode_tekst_to[i] == 82:
                        btype_beredskap_areal += bebygdareal[i]
                        btype_beredskap_antall += 1
            

            #BEREGNE TALKKAREAL
            #Skråtak
            print(kom_nr)
            if btype_enebolig_antall == 0:
                print("Her er det null :S")
            else:
                print("Tut og kjør")

            # tilgjengelig_takareal_enebolig = btype_enebolig_areal * utnyttingsgrad_skratak_enebolig
            
            # takareal_enebolig = ((math.sqrt((btype_enebolig_areal / btype_enebolig_antall) / (bygningsmodell_smahus_lengde / bygningsmodell_smahus_bredde)) / 2) / (math.cos(math.radians(takvinkel_enebolig))) * (math.sqrt((btype_enebolig_areal / btype_enebolig_antall) / (bygningsmodell_smahus_lengde / bygningsmodell_smahus_bredde)) * (bygningsmodell_smahus_lengde / bygningsmodell_smahus_bredde)) * 2) * btype_enebolig_antall
            

            
            # print("Areak:", btype_enebolig_areal)
            # print("Antall:", btype_enebolig_antall)
            #print(math.sqrt((btype_enebolig_areal / btype_enebolig_antall)))
            print("")


                #### SKRIVER TIL REGNEARK!!!!!!!!!!!!!!
    #         skriver = wb['Fylke ' + str(f+1)] #Setter excel-ark som skal skrives til
    #         skriver.cell(1,(2+row)).value = kom_navn
    #         skriver.cell(2,(2+row)).value = btype_enebolig   #cell[ned, venstre]
    #         skriver.cell(3,(2+row)).value = btype_tomannsbolig
    #         skriver.cell(4,(2+row)).value = btype_rekkehus
    #         skriver.cell(5,(2+row)).value = btype_storeboliger
    #         skriver.cell(6,(2+row)).value = btype_bofellesskap
    #         skriver.cell(7,(2+row)).value = btype_fritidsbolig
    #         skriver.cell(8,(2+row)).value = btype_koie
    #         skriver.cell(9,(2+row)).value = btype_garasje
    #         skriver.cell(10,(2+row)).value = btype_annenbolig
    #         skriver.cell(11,(2+row)).value = btype_industri
    #         skriver.cell(12,(2+row)).value = btype_lager
    #         skriver.cell(13,(2+row)).value = btype_fiskeri
    #         skriver.cell(14,(2+row)).value = btype_kontor
    #         skriver.cell(15,(2+row)).value = btype_forretning
    #         skriver.cell(16,(2+row)).value = btype_messekongress

    # wb.save(regneark_eksisterende) #Lagrer excel-fil

print("\nferdig")