
#importere nødvendige bibliotek
import pandas as pd
import requests

#Plassering av excel-dokument
kildedata = pd.read_excel("C:/Prosjekter/NELFO_solcelle/fme_excel_export.xlsx")

antall = len(kildedata) #Antall rader i exceldokument.

#Setter alle varialbler
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

#Henter alle kommunenummer fra kommunenummerregister i Geonorge (Kartverket) som kan itereres over
kom_url = "https://register.geonorge.no/api/sosi-kodelister/kommunenummer.json?" #URL til kodeliste
r = requests.get(kom_url) #Kobler opp til url
kom_data = r.json() #henter innhold som JSON

kom_antall = len(kom_data['containeditems']) #henter alle kommuner i kodelisten

for kom in range(kom_antall): #Looper gjennom kodelisten f0r hver kommune
    kom_nr = kom_data['containeditems'][kom]['codevalue'] #kommunenummer
    if kom_antall != 2100 or kom_antall != 2211: #Utelukker Svalbard og Jan Mayen
        print(kom_nr)

# for i in range(antall):
#     if kommunenummer[i] == 4222:
#         print("Treff")