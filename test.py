# import pandas as pd
# import requests, openpyxl, math


# #HENTER INN MELLOMREGNING FRA EXCELARK FOR MELLOMREGNING
# mellomregning_kwh = pd.read_excel("C:/Prosjekter/NELFO_solcelle/excel_calculations/calc_and_inputxlsx.xlsx", "kWh")

# kwh_tak_enebolig = mellomregning_kwh['Solstrøm tak per bygg [kWh/år]'][0]

# print("Kwh tak:", kwh_tak_enebolig)


# import requests

# fylke_url = "https://register.geonorge.no/api/sosi-kodelister/fylkesnummer.json?"
# r_fylke = requests.get(fylke_url)
# fylke_data = r_fylke.json()
# fylke_antall = len(fylke_data['containeditems'])

# fylke = [] #Liste som populeres med alle fylkesnummer utenom Svalbard og Jan Mayen

# for fylkekode in range(fylke_antall): #Loop som looper gjennom fylke_url og henter alle fylkesnummer
#     fylkenr = fylke_data['containeditems'][fylkekode]['codevalue'] #Lokasjon for fylkesnummer i json-fil (fylke_url)
#     if fylkenr != "21" and fylkenr != "22": #Utelukker Svalbard og Jan Mayen
#         fylke.append(fylkenr)

# fylke.sort() #Sorterer fylkeslisten slik at minste fylkesnummer kommer først.

# for f in range(len(fylke)):
#     if str(fylke[f]) == "03":
#         fylkesnavn = "Oslo"
#     elif str(fylke[f]) == "11":
#         fylkesnavn = "Rogaland"
#     elif str(fylke[f]) == "15":
#         fylkesnavn = "Møre og Romsdal"
#     elif str(fylke[f]) == "18":
#         fylkesnavn = "Nordland"
#     elif str(fylke[f]) == "30":
#         fylkesnavn = "Viken"
#     elif str(fylke[f]) == "34":
#         fylkesnavn = "Innlandet"
#     elif str(fylke[f]) == "38":
#         fylkesnavn = "Vestfold og Telemark"
#     elif str(fylke[f]) == "42":
#         fylkesnavn = "Agder"
#     elif str(fylke[f]) == "46":
#         fylkesnavn = "Vestland"
#     elif str(fylke[f]) == "50":
#         fylkesnavn = "Trøndelag"
#     elif str(fylke[f]) == "54":
#         fylkesnavn = "Troms og Finnmark"

#     print(fylkesnavn) 
#     print(fylke[f])
#     print()

# print(fylke)

# import requests

# fylke_url = "https://register.geonorge.no/api/sosi-kodelister/fylkesnummer.json?"
# r_fylke = requests.get(fylke_url)
# fylke_data = r_fylke.json()
# fylke_antall = len(fylke_data['containeditems'])

# fylke = [] #Liste som populeres med alle fylkesnummer utenom Svalbard og Jan Mayen

# for fylkekode in range(fylke_antall): #Loop som looper gjennom fylke_url og henter alle fylkesnummer
#     fylkenr = fylke_data['containeditems'][fylkekode]['codevalue'] #Lokasjon for fylkesnummer i json-fil (fylke_url)
#     if fylkenr != "21" and fylkenr != "22": #Utelukker Svalbard og Jan Mayen
#         fylke.append(fylkenr)

# fylke.sort() #Sorterer fylkeslisten slik at minste fylkesnummer kommer først.

# print(fylke)

# kom_url = "https://register.geonorge.no/api/sosi-kodelister/kommunenummer.json?" #URL til kodeliste
# r = requests.get(kom_url) #Kobler opp til url
# kom_data = r.json() #henter innhold som JSON

# kom_antall = len(kom_data['containeditems']) #henter alle kommuner i kodelisten

# fylke = ["03"]

# for f in range(len(fylke)): #Looper over alle fylkene
#     print("Prosesserer fylke:", fylke[f])
#     row = 0 #Teller som teller hvert fylke. 
#     for kom in range(kom_antall): #Looper gjennom kodelisten for hver kommune
#         kom_nr = kom_data['containeditems'][kom]['codevalue'] #kommunenummer
#         kom_navn = kom_data['containeditems'][kom]['label']
#         if str(fylke[f]) == str(kom_nr[:2]): #Velger kommuner i gitt fylke
#             row += 1

#             print(type(kom_nr), kom_nr)

#             print(input("input"))

# print("ferdig")

#importere nødvendige bibliotek
import pandas as pd
import requests, openpyxl, math, shutil
import xlwings as xl
import geopandas as gpd

kildedata = gpd.read_file("C:/Prosjekter/NELFO_solcelle/Bygninger_kom_fylke.gpkg")


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

antall = len(kildedata) 

for i in range(antall):
    print (kommunenummer[i])
    if str(kommunenummer[i]) == "301":
        print("JOHANNES HER ER OSLO!!!!!!!!!!!!!!!!!!!!!")
        import time
        time.sleep(10)
