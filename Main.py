# Données techniques Verdun


import pyperclip
import os
import sys
import pandas as pd
import numpy as np
from Mains import mains
from contextlib import redirect_stdout
from Ctes import directory, dir2, dir_csv, dir_log, dir_anc_codes, dir_ld

sys.stdout = open(dir_log,"w")

Source_Workbook = []

for filename in os.listdir(directory):
    if filename.endswith(".xls") or filename.endswith(".XLS"):
        Source_Workbook.append(filename)

liste_csv, liste_noire, liste_blanche, liste_vari, log_exceptions, anc_codes, codes = mains(Source_Workbook)

np.savetxt(dir_csv,
           liste_csv,
           delimiter=";",
           fmt='% s',
           encoding='UTF-8')

# np.savetxt(dir_anc_codes,
#            anc_codes,
#            delimiter=";",
#            fmt='% s',
#            encoding='UTF-8')

# # print(liste_noire)
# tablenoire = tableize(pd.DataFrame(liste_noire, columns=["Fichiers"]))
# pyperclip.copy(tablenoire)
# print(tablenoire)
affichage = [liste_noire, liste_blanche, liste_vari, log_exceptions, codes]
# affichage = [liste_vari]
for liste in affichage:
#     print(len(liste)) #(f'{liste=}'.split('=')[0],
#     for el in liste:
#         print(el)
#         print("\n")
    print(liste)
# # print(table_anc_code)
# # t = pd.DataFrame(anc_codes, index = [0])
# # t.to_excel(dir_anc_codes)
# # ld = pd.DataFrame(liste_noire, index = [0])
# # ld.to_excel(dir_ld)
# # f.close()
# # print(table_anc_code)
# chara = "["
# for upl in liste_blanche:
#     chara += f'{upl};'

# pyperclip.copy(chara)
sys.stdout.close()



