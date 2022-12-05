#importere nødvendige bibliotek
import pandas as pd
import requests, openpyxl, math

produksjonstall_excel = pd.read_excel("C:/Prosjekter/NELFO_solcelle/input_fag/produksjonstall_kommune.xlsx")

produksjonstall_kommunenummer = produksjonstall_excel['Kommunenummer']
kwh_south = produksjonstall_excel['kwh/m2_south_f']
kwh_east = produksjonstall_excel['kwh/m2_east_f']
kwh_west = produksjonstall_excel['kwh/m2_west_f']
kwh_flat = produksjonstall_excel['kwh/m2_flat_r']
kwh_25south = produksjonstall_excel['kwh/m2_25south_r']
kwh_25east = produksjonstall_excel['kwh/m2_25east_r']
kwh_25west = produksjonstall_excel['k2w/m2_25west_r']


kwp_south = produksjonstall_excel['kwh/kwp_south_f']
kwp_east = produksjonstall_excel['kwh/kwp_east_f']
kwp_west = produksjonstall_excel['kwh/kwp_west_f']
kwp_flat = produksjonstall_excel['kwh/kwp_flat_r']
kwp_25south = produksjonstall_excel['kwh/kwp_25south_r']
kwp_25east = produksjonstall_excel['kwh/kwp_25east_r']
kwp_25east = produksjonstall_excel['k2w/kwp_25west_r']

installert_south = produksjonstall_excel['installert_south_f']
installert_east = produksjonstall_excel['innstallert_east_f']
installert_west = produksjonstall_excel['installert_west_f']
installert_flat = produksjonstall_excel['installert_flat_r']
installert_25south = produksjonstall_excel['installert_25south']
installert_25east = produksjonstall_excel['installert_25east']
installert_25west = produksjonstall_excel['installert_25west']



# print(kwh_south[292])
# print(kwh_east[292])
# print(kwh_west[292])
# print(kwh_flat[292])
# print(kwh_25south[292])
# print(kwh_25east[292])
# print(kwh_25west[292])
# print(kwp_south[292])
# print(kwp_east[292])
# print(kwp_west[292])
# print(kwp_flat[292])
# print(kwp_25south[292])
# print(kwp_25east[292])
# print(kwp_25east[292])
# print(installert_south[292])
# print(installert_east[292])
# print(installert_west[292])
# print(installert_flat[292])
# print(installert_25south[292])
# print(installert_25east[292])
# print(installert_25west[292])

for kommuneproduksjon in range(len(produksjonstall_kommunenummer)):
    print(produksjonstall_kommunenummer[kommuneproduksjon])