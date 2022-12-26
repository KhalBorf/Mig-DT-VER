# Ceci nous permet de créer une liste à partir d'un Excel
import pandas as pd
from Objects2 import CoordonnéesExcel, FindExcel, tableize
import pyperclip
file = r'C:\Users\k.elmaliki\OneDrive - BORFLEX BORFLEX CTR-035961\Bureau\Anciens codes ++.xlsx'
wb = pd.read_excel(file, index_col=0, header = None)
données = CoordonnéesExcel(wb, 'end', 1, "B", "b", "C")
liste = données.get_tout()
dictio = {}
dictio2 = '{'
i = 0
for tuple in liste:
    dictio[tuple[0]] = tuple[1]
    dictio2 += f'\'{tuple[0]}\': \'{tuple[1]}\','
print(dictio)
dictio2 += '}'
print(dictio2)
pyperclip.copy(dictio2)

