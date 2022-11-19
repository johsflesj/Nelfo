#Programmet leser excel-fil som er eksport av GIS-data. Denne filen heter:
#Programmet henter eksisterende excel-mal og skriver til ny fil.

# OBS! Programmet må oppdateres med riktig lenke til filer som programmet leser og skriver til.

#Biblioteker som må installeres: Pandas, requests, xlsxwriter, openpyxl


#importere nødvendige bibliotek
import pandas as pd
import requests, openpyxl

#Excel-fil som inneholder GIS-uttrekk
kildedata = pd.read_excel("C:/Prosjekter/NELFO_solcelle/fme_excel_export.xlsx")

#Excel-fil vi skal skrive data til:
regneark_eksisterende = "C:/Prosjekter/NELFO_solcelle/Mal_test.xlsx"
wb = openpyxl.load_workbook(regneark_eksisterende)

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

fylke = ["03", 11, 15, 18, 30, 34, 38, 42, 46, 50, 54]

for f in range(1): #!!!!! MÅ BYTTES TIL len(fylke)
    for kom in range(kom_antall): #Looper gjennom kodelisten for hver kommune
        kom_nr = kom_data['containeditems'][kom]['codevalue'] #kommunenummer
        if str(fylke[f]) == str(kom_nr[:2]):

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

            #Bygnignstyper: Sette bruke bygningtypenummer i navn?
            btype_enebolig = 0
            btype_tomannsbolig = 0
            btype_rekkehus = 0
            btype_storeboliger = 0
            btype_bofellesskap = 0
            btype_fritidsbolig = 0
            btype_koie = 0
            btype_garasje = 0
            btype_annenbolig = 0
            btype_industri = 0
            btype_lager = 0
            btype_fiskeri = 0
            btype_kontor = 0
            btype_forretning = 0
            btype_messekongress = 0
            btype_terminal = 0
            btype_garasjehangar = 0
            btype_vegtrafikk = 0
            btype_hotell = 0
            btype_overnatting = 0
            btype_restaurant = 0
            btype_skole = 0
            btype_universitet = 0
            btype_museum = 0
            btype_idrett = 0
            btype_kulturhus = 0
            btype_religios = 0
            btype_sykehus = 0
            btype_sykehjem = 0
            btype_primarhelse = 0
            btype_beredskap = 0




                
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
                    bygningstypekode_tekst_tre_tot += bygningstypekode_tekst_tre[i]
                    tellekolonne_tot += tellekolonne[i]
                    bygningskode_tekst_to_tot += bygningskode_tekst_to[i]
                    shape_length_tot += shape_length[i]
                    shape_area_tot += shape_area[i]
            print(bebygdareal_tot)
            print("")


            # with pd.ExcelWriter("C:/Prosjekter/NELFO_solcelle/tom_test.xlsx", mode="a", engine="openpyxl", if_sheet_exists='overlay') as writer:
            #     test = pd.DataFrame([fkbareal_tot])
            #     test.to_excel(writer, sheet_name="Fylke 1", startrow=1, startcol=1, index=False, header=False)  

            test = wb['Fylke 1'] #Setter excel-ark som skal skrives til
            test.cell(2,3).value = "test123"   #cell[ned, venstre]

    wb.save(regneark_eksisterende) #Lagrer excel-fil

print("\nferdig")