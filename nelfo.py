#Programmet leser excel-fil som er eksport av GIS-data. Denne filen heter:
#Programmet henter eksisterende excel-mal og skriver til ny fil.


#importere nødvendige bibliotek
import pandas as pd
import requests

#Plassering av excel-dokument
kildedata = pd.read_excel("C:/Prosjekter/NELFO_solcelle/fme_excel_export.xlsx")

antall = len(kildedata) #Antall rader i exceldokument.

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

#Henter alle kommunenummer fra kommunenummerregister i Geonorge (Kartverket) som kan itereres over
kom_url = "https://register.geonorge.no/api/sosi-kodelister/kommunenummer.json?" #URL til kodeliste
r = requests.get(kom_url) #Kobler opp til url
kom_data = r.json() #henter innhold som JSON

kom_antall = len(kom_data['containeditems']) #henter alle kommuner i kodelisten

for kom in range(kom_antall): #Looper gjennom kodelisten for hver kommune
    kom_nr = kom_data['containeditems'][kom]['codevalue'] #kommunenummer
    if int(kom_nr) == 5001 or int(kom_nr) == 1865: ##!= 2100 or str(kom_nr) != 2211: #Utelukker Svalbard og Jan Mayen
        print(kom_nr)

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
        elspotomr_tot = 0 
        bygningstypekode_tekst_tre_tot = 0 
        tellekolonne_tot = 0 
        bygningskode_tekst_to_tot = 0 
        name_tot = 0 
        shape_length_tot = 0 
        shape_area_tot = 0 

            
        for i in range(antall):
            if int(kom_nr) == int(kommunenummer[i]):
                #print(bygningstypekode_tekst_tre[i])
                bebygdareal_tot += bebygdareal[i]
                fkbareal_tot += fkbareal[i]
                bruksarealtilbolig_tot += bruksarealtilbolig[i]
                bruksarealtilannet_tot += bruksarealtilannet[i]
                alternativtareal_tot += alternativtareal[i]
                alternativtareal2_tot += alternativtareal2[i]
                utenbebygdareal_tot += utenbebygdareal[i]
                antalletasjer_tot += antalletasjer[i]
                elspotomr_tot += elspotomr_tot[i]
                bygningstypekode_tekst_tre_tot += bygningstypekode_tekst_tre[i]
                tellekolonne_tot += tellekolonne[i]
                bygningskode_tekst_to_tot += bygningskode_tekst_to[i]
                name_tot += name[i]
                shape_length_tot += shape_length[i]
                shape_area_tot += shape_area[i]
        print(bebygdareal_tot)
        print("")


            #excelskriver = pd.ExcelWriter('Test.xlsx, mode="a", engine="auto", ')

print("\nferdig")