try:
    (colonne_temps, ligne_1) = var1.fetch_colonne_ligne("nr", "B", "choix")
    colonne_1 = 1
    liste_blanche.append(book)
except Exception as e:
    try:
        (colonne_temps, ligne_1) = var1.fetch_colonne_ligne("nr", "A", "choix")
        liste_blanche.append(book)
        colonne_1 = 0
    except Exception as e:
        liste_noire.append((book, "Pas de choix"))
        continue

try:  ##################################################### Code
    colonne_anc_code = var1.find_colonne("code", ligne_1)
except Exception as e:
    liste_vari.append(("Pas d\'ancien code", book))
    continue

try:  ##################################################### Sylob
    colonne_code_Sylob = var1.find_colonne("Sylob", ligne_1)
    liste_sylob.append(book)
    Sylob = True
except Exception as e:
    pass

try:  ##################################################### Produit
    colonne_produit = var1.find_colonne("produit", ligne_1)
except Exception as e:
    liste_vari.append(("Pas de produit", book))
    continue

try:  ##################################################### Total
    ligne_total = var1.find_ligne("total", colonne_produit)
except Exception as e:
    liste_vari.append(("Pas de poids total", book))
    continue

try:  ##################################################### Poids
    colonne_poids = var1.find_colonne("Kg", ligne_1)
except Exception as e:
    liste_vari.append(("Pas de titre poids", book))
    continue

try:  ##################################################### Temps
    ligne_temps = var1.find_ligne("temps", colonne_temps)
except Exception as e:
    liste_vari.append(("Pas de temps", book))
    continue

try:  ##################################################### Opération
    colonne_opération = var1.find_colonne("opération", ligne_temps)
except Exception as e:
    liste_vari.append(("Pas d\'opération ", book))
    continue

try:  ##################################################### Vitesse
    colonne_vitesse = var1.find_colonne("vites", ligne_temps)
except Exception as e:
    liste_vari.append(("Pas de vitesse", book))
    continue

try:  ##################################################### BT5
    colonne_bt5 = var1.find_colonne("bt5", ligne_temps)
except Exception as e:
    liste_vari.append(("Pas de bt5", book))
    continue

try:
    ligne_temps2 = var1.find_ligne("Base partie", colonne_produit)
except Exception as e:
    print((book, "pas de base p2"))
    continue