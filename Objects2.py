import pandas as pd

class CoordonnéesExcel:

    def __init__(self, fichier, espace, initial, test, *args):
        self.celulles = []
        for cell in args:
            if type(cell).__name__ == "str":
                cell = ord(cell.lower()) - 97
                self.celulles.append(cell)
            else:
                cell = ord(cell[0].lower()) - 97 + cell[1]
                self.celulles.append(cell)
        self.fichier = fichier
        self.espace = espace
        self.liste = []
        self.itération = -1
        self.initial = initial
        self.test = ord(test.lower()) - 97
        i = self.initial - 1
        if type(self.espace).__name__ == "int":
            n = 0
            while n != self.espace:
                if str(self.fichier.iloc[i][self.test]) != 'nan':
                    pliste = [str(self.fichier.iloc[i][cell]) for cell in self.celulles]
                    self.liste.append(pliste)
                    n = 0
                else:
                    n += 1
                i += 1
        elif type(self.espace).__name__ == "str":
            while not(self.espace.lower() in str(self.fichier.iloc[i][self.test]).lower()):
                if str(self.fichier.iloc[i][self.test]) != 'nan':
                    pliste = [str(self.fichier.iloc[i][cell]) for cell in self.celulles]
                    self.liste.append(pliste)
                i += 1
    def get_tout(self):
        return self.liste

class CoordonnéesExcelGamme:

    def __init__(self, fichier, espace, initial, test, *args):
        self.celulles = []
        for cell in args:
            if type(cell).__name__ == "str":
                cell = ord(cell.lower()) - 97
                self.celulles.append(cell)
            else:
                cell = ord(cell[0].lower()) - 97 + cell[1]
                self.celulles.append(cell)
        self.fichier = fichier
        self.espace = espace
        self.liste = []
        self.itération = -1
        self.initial = initial
        self.test = ord(test.lower()) - 97
        i = self.initial - 1
        if type(self.espace).__name__ == "int":
            n = 0
            while n != self.espace:
                if str(self.fichier.iloc[i][self.test]) != 'nan':
                    pliste = [str(self.fichier.iloc[i][cell]) for cell in self.celulles]
                    self.liste.append(pliste)
                    n = 0
                else:
                    n += 1
                i += 1
        elif type(self.espace).__name__ == "str":
            n = 0
            liste_vide = [" " for cell in self.celulles]
            # liste_vide.pop(len(liste_vide) - 1)

            while not(self.espace.lower() in str(self.fichier.iloc[i][self.test]).lower()):
                if str(self.fichier.iloc[i][self.test]) == 'nan':
                    pliste = liste_vide
                    # pliste = pliste + ["<cr/>"]
                    n +=1
                else:
                    pliste = [str(self.fichier.iloc[i][cell]) for cell in self.celulles]
                    n=0
                i += 1
                self.liste.append(pliste)
            i = 0
            if n != 0:
                for i in range(n):
                    self.liste.pop(len(self.liste)-1)
    def get_tout(self):
        return self.liste

class CoordonnéesExcelGamme2Temps:

    def __init__(self, fichier, espace, initial, test, *args):
        self.celulles = []
        for cell in args:
            if type(cell).__name__ == "str":
                cell = ord(cell.lower()) - 97
                self.celulles.append(cell)
            else:
                cell = ord(cell[0].lower()) - 97 + cell[1]
                self.celulles.append(cell)
        self.fichier = fichier
        self.espace = espace
        self.liste = []
        self.itération = -1
        self.initial = initial
        self.test = ord(test.lower()) - 97
        i = self.initial - 1
        if type(self.espace).__name__ == "int":
            n = 0
            while n != self.espace:
                if str(self.fichier.iloc[i][self.test]) != 'nan':
                    pliste = [str(self.fichier.iloc[i][cell]) for cell in self.celulles]
                    self.liste.append(pliste)
                    n = 0
                else:
                    n += 1
                i += 1
        elif type(self.espace).__name__ == "str":
            n = 0
            liste_vide = [" " for cell in self.celulles]
            # liste_vide.pop(len(liste_vide) - 1)
            # liste_nan = ['nan' for cell in self.celulles]

            while not(self.espace.lower() in str(self.fichier.iloc[i][self.test]).lower()):
                pliste = [" " if str(self.fichier.iloc[i][cell]) == 'nan' else str(self.fichier.iloc[i][cell]) for cell in self.celulles]
                if pliste == liste_vide:
                    # pliste = liste_vide
                    # pliste = pliste + ["<cr/>"]
                    n +=1
                else:
                    # pliste = [str(self.fichier.iloc[i][cell]) for cell in self.celulles]
                    n=0
                i += 1
                self.liste.append(pliste)
            i = 0
            if n != 0:
                for i in range(n):
                    self.liste.pop(len(self.liste)-1)
    def get_tout(self):
        return self.liste

    def get_colonnes(self):
        liste_vide = [" " for cell in self.celulles]
        return liste_vide

class FindExcel:

    def __init__(self, fichier):
        self.fichier = fichier

    def find_ligne(self, char, colonne):
        colonne = ord(colonne.lower()) - 97
        if colonne > 0:
            i = 0
            while True:
                i += 1
                if char.lower() in str(self.fichier.iloc[i][colonne]).lower():
                    break
            return i+1
        else:
            index = self.fichier.index
            i = 0
            for ind in index:
                i += 1
                if char.lower() in str(ind).lower():
                    return i
            return "Recherche non aboutit"

    def find_ligne_n(self, char, colonne, n):
        colonne = ord(colonne.lower()) - 97
        if colonne > 0:
            i = 0
            j = 0
            while True:
                i += 1
                if j >= n:
                    if char.lower() in str(self.fichier.iloc[i][colonne]).lower():
                        break
                else:
                    if char.lower() in str(self.fichier.iloc[i][colonne]).lower():
                        j += 1
            return i+1
        else:
            index = self.fichier.index
            i = 0
            for ind in index:
                i += 1
                if char.lower() in str(ind).lower():
                    return i
            return "Recherche non aboutit"

    def find_colonne(self, char, ligne):
        ligne -= 1
        i = 0
        while True:
            i += 1
            if char.lower() in str(self.fichier.iloc[ligne][i]).lower():
                break
        return chr(i + 97)

    def fetch_value(self, char, colonne, target_colonne):
        target_colonne = ord(target_colonne.lower()) - 97
        return str(self.fichier.iloc[int(self.find_ligne(char, colonne))][int(target_colonne)])

    def fetch_ligne(self):
        pass

    def fetch_colonne(self, char, colonne, target_value):
        return self.find_colonne(target_value, self.find_ligne(char, colonne))

    def fetch_colonne_ligne(self, char, colonne, target_value):
        return self.fetch_colonne(char, colonne, target_value), self.find_ligne(char, colonne)

def tableize(df):
    if not isinstance(df, pd.DataFrame):
        return
    df_columns = df.columns.tolist()
    max_len_in_lst = lambda lst: len(sorted(lst, reverse=True, key=len)[0])
    align_center = lambda st, sz: "{0}{1}{0}".format(" "*(1+(sz-len(st))//2), st)[:sz] if len(st) < sz else st
    align_right = lambda st, sz: "{0}{1} ".format(" "*(sz-len(st)-1), st) if len(st) < sz else st
    max_col_len = max_len_in_lst(df_columns)
    max_val_len_for_col = dict([(col, max_len_in_lst(df.iloc[:,idx].astype('str'))) for idx, col in enumerate(df_columns)])
    col_sizes = dict([(col, 2 + max(max_val_len_for_col.get(col, 0), max_col_len)) for col in df_columns])
    build_hline = lambda row: '+'.join(['-' * col_sizes[col] for col in row]).join(['+', '+'])
    build_data = lambda row, align: "|".join([align(str(val), col_sizes[df_columns[idx]]) for idx, val in enumerate(row)]).join(['|', '|'])
    hline = build_hline(df_columns)
    out = [hline, build_data(df_columns, align_center), hline]
    for _, row in df.iterrows():
        out.append(build_data(row.tolist(), align_right))
    out.append(hline)
    return "<cr/>".join(out) #<\n>
def tableize2(df):
    if not isinstance(df, pd.DataFrame):
        return
    df_columns = df.columns.tolist()
    max_len_in_lst = lambda lst: len(sorted(lst, reverse=True, key=len)[0])
    align_center = lambda st, sz: "{0}{1}{0}".format(" "*(1+(sz-len(st))//2), st)[:sz] if len(st) < sz else st
    align_right = lambda st, sz: "{0}{1} ".format(" "*(sz-len(st)-1), st) if len(st) < sz else st
    max_col_len = max_len_in_lst(df_columns)
    max_val_len_for_col = dict([(col, max_len_in_lst(df.iloc[:,idx].astype('str'))) for idx, col in enumerate(df_columns)])
    col_sizes = dict([(col, 2 + max(max_val_len_for_col.get(col, 0), max_col_len)) for col in df_columns])
    build_hline = lambda row: ''.join(['' * col_sizes[col] for col in row]).join(['', ''])
    build_data = lambda row, align: " ".join([align(str(val), col_sizes[df_columns[idx]]) for idx, val in enumerate(row)]).join(['', ''])
    hline = build_hline(df_columns)
    out = [hline, build_data(df_columns, align_center), hline]
    for _, row in df.iterrows():
        out.append(build_data(row.tolist(), align_right))
    out.append(hline)
    return "<cr/>".join(out) #<\n>

