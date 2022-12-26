import pandas as pd
from Objects2 import CoordonnéesExcel, CoordonnéesExcelGamme, CoordonnéesExcelGamme2Temps, FindExcel, tableize, tableize2
from Ctes import EDI_article_dict, EDI_opération_dict, EDI_composants_dict, directory, codeDT, libDT, liste_noire_finale, liste_composants_Sylob, liste_gris_finale


def mains (Source_Workbook):
    class Poids_nulle(Exception):
        pass
    dictio_ouvert = {}
    dictio_centre1 = {}
    dictio_centre2 = {}
    liste_noire = []
    liste_blanche = []
    liste_sylob = []
    liste_vari = []
    liste_csv = []
    log_exceptions = []
    log_sans_code_Sylob = []
    table_anc_code = dict({})
    table_codes_articles = {"131100v-0": "A03-00385"}
    centre_de_charge_dosage = "BUV-DOS-V"
    centre_de_charge_melangeage = "BUV-MEL-V"
    centre_de_charge_acceleration = "BUV-MEL-O"
    # yes_ = False
    # book2 = ""
    for book in Source_Workbook:
        # if yes_ != True:
        #     liste_noire.append(book2)
        # yes_ = False
        temps2 = False
        book2 = book
        book1 = book
        EDI_article_liste =[]
        EDI_opération_liste = []
        EDI_composants_liste = []
        Sylob = False
        colonne_anc_code = ""
        colonne_code_Sylob = ""
        colonne_produit = ""
        colonne_poids = ""
        ligne_total = 0
        colonne_phr = ""
        colonne_temps = ""
        colonne_opération = ""
        colonne_vitesse = ""
        colonne_bt5 = ""
        poids_total = ""
        premiere_colonne = 0
        ligne_1 = 0
        ligne_temps2 = 0
        poids_total2 = 0
        temps_OP = '0'
        temps_OP_2 = '0'



        lieux_stockage_mp = "8ac118828068f94a01806bcae0b90b8d"
        EDI_composants_dict["CODE_LIEU"] = lieux_stockage_mp
        # EDI_opération_dict["Opé_Centre"] = centre_de_charge_dosage
        wb = pd.read_excel(directory + '\\' + book, index_col=0, header = None)
        var1 = FindExcel(wb)

        # if book.lower() != '383110v.xls':
        #     continue

        # if (book in liste_noire_finale) or (book in liste_gris_finale):
        #     continue

        if book1.endswith(".xls"):
            book1 = book1.replace(".xls", "")
        else:
            book1 = book1.replace(".XLS", "")
        ########################################################################################################################
        # Puissque les positions des cellules dans les fichiers Excel ne sont pas forcément identiques,
        # en se basant sur leurs contenu, le bloc ci-dessous sert à repérer les colonnes et/ou lignes de ces cellules
        ########################################################################################################################
        if True:
            try:##################################################### Ligne Choix Du Mélangeur
                (colonne_temps, ligne_1) = var1.fetch_colonne_ligne("nr", "B", "choix")
                colonne_1 = 1

            except Exception as e:
                try:
                    (colonne_temps, ligne_1) = var1.fetch_colonne_ligne("nr", "A", "choix")

                    colonne_1 = 0
                except Exception as e:
                    liste_noire.append((book, "Pas de choix"))
                    continue

            try:##################################################### Colonne Ancien Code
                colonne_anc_code = var1.find_colonne("code", ligne_1)
            except Exception as e:
                liste_vari.append(("Pas d\'ancien code", book))
                continue

            try:##################################################### Colonne Code Sylob
                colonne_code_Sylob = var1.find_colonne("Sylob", ligne_1)
                liste_sylob.append(book)
                Sylob = True
            except Exception as e:
                pass

            try:##################################################### Colonne Produit MP
                colonne_produit = var1.find_colonne("produit", ligne_1)
            except Exception as e:
                liste_vari.append(("Pas de produit", book))
                continue

            try:##################################################### Ligne Poids Total
               ligne_total = var1.find_ligne("total", colonne_produit)
            except Exception as e:
                liste_vari.append(("Pas de poids total", book))
                continue

            try:##################################################### Colonne Poids Total
                colonne_poids = var1.find_colonne("Kg", ligne_1)
            except Exception as e:
                liste_vari.append(("Pas de titre poids", book))
                continue

            try:##################################################### Ligne Temps Gamme
                ligne_temps = var1.find_ligne("temps", colonne_temps)
            except Exception as e:
                liste_vari.append(("Pas de temps", book))
                continue

            try:##################################################### Colonne Opération
                colonne_opération = var1.find_colonne("opération", ligne_temps)
            except Exception as e:
                liste_vari.append(("Pas d\'opération ", book))
                continue

            try:##################################################### Colonne Vitesse
                colonne_vitesse = var1.find_colonne("vites", ligne_temps)
            except Exception as e:
                liste_vari.append(("Pas de vitesse", book))
                continue

            try:##################################################### Colonne BT5
                colonne_bt5 = var1.find_colonne("bt5", ligne_temps)
            except Exception as e:
                liste_vari.append(("Pas de bt5", book))
                continue

            try:##################################################### Ligne Base partie 2
                ligne_temps2 = var1.find_ligne("Base partie", colonne_produit)
            except Exception as e:
                liste_vari.append((book, "pas de base p2"))
                continue

            try:##################################################### Ligne Mélange complet
                ligne_melange_comp = var1.find_ligne("Mélange complet", colonne_produit)
            except Exception as e:
                liste_vari.append((book, "pas de Mélange complet"))
                continue

            try:##################################################### Ligne Temps OP
                ligne_temps_OP = var1.find_ligne("Temps (s)", colonne_opération)
            except Exception as e:
                liste_vari.append((book, "Pas de ligne temps OP"))

            try:##################################################### Ligne Temps OP 2
                ligne_temps_OP_2 = var1.find_ligne_n("Temps (s)", colonne_opération, 1)
            except Exception as e:
                liste_vari.append((book, "Pas de ligne temps OP 2"))


            try:##################################################### Test des cases poids: si le poids de melange complet est nul on test le poids d'un temps. Si ce dernier est nul on ne continue pas
                poids_total = str(wb.iloc[int(ligne_melange_comp) - 1][ord(colonne_poids.lower()) - 97])
                if (poids_total != "nan"):
                    temps2 = True
                else:
                    poids_total = str(wb.iloc[int(ligne_total) - 1][ord(colonne_poids.lower()) - 97])
                    if (poids_total == "nan"):
                        raise Poids_nulle()
            except Exception as e:
                log_exceptions.append((book, e))
                continue

            try:##################################################### Test de la case Temps OP: si == rien on le met à '0'. le test des cases de temps n'est pas bloquant
                temps_OP = str(wb.iloc[int(ligne_temps_OP)][ord(colonne_opération.lower()) - 97])
                if temps_OP == "nan":
                    log_exceptions.append((book, "Temps OP null"))
                    raise Exception
            except Exception as e:
                log_exceptions.append((book, "Erreur temps OP"))
                temps_OP = '0'

            if temps2: ##################################################### Test de la case Temps OP 2: si == rien on le met à '0'. le test des cases de temps n'est pas bloquant
                try:
                    temps_OP_2 = str(wb.iloc[int(ligne_temps_OP_2)][ord(colonne_opération.lower()) - 97])
                    if temps_OP_2 == "nan":
                        log_exceptions.append((book, "Temps OP 2 null"))
                except Exception as e:
                    log_exceptions.append((book, "Erreur temps OP"))
                    temps_OP_2 = '0'

        ########################################################################################################################
        # if not (centre1 in dictio_centre1):
        #     dictio_centre1[centre1] = centre2
        # if not (centre2 in dictio_centre2):
        #     dictio_centre2[centre2] = centre1


        ########################################################################################################################
        # S'il existe une colonne Sylob dans le fichier (nouveau modèle) on extrait les codes Sylob depuis le fichier. Sinon, on n'extrait que les anciens codes
        ########################################################################################################################
        if Sylob:
            Objet_composants = CoordonnéesExcel(wb, 'Total', ligne_1 + 1, colonne_produit, colonne_code_Sylob, colonne_poids)
            if temps2:
                Objet_composants_2temps = CoordonnéesExcel(wb, 'Mélange complet', ligne_temps2 + 1, colonne_produit, colonne_code_Sylob, colonne_poids)

            # colonne_name = "Code Article"
        else:
            if temps2:
                Objet_composants_2temps = CoordonnéesExcel(wb, 'Mélange complet', ligne_temps2 + 1, colonne_produit, colonne_anc_code, colonne_poids)

            Objet_composants = CoordonnéesExcel(wb, 'Total', ligne_1 + 1, colonne_produit, colonne_anc_code, colonne_poids)

            # colonne_name = "Ancien code"

        ########################################################################################################################
        # On extrait les données suivantes: les composants(code, désignation, poids), gamme
        ########################################################################################################################
        Objet_ancien_code = CoordonnéesExcel(wb, 'Total', ligne_1 + 1, colonne_produit, colonne_produit, colonne_anc_code)
        Objet_mode_operatoire_1temps = CoordonnéesExcelGamme(wb, 'Densité', ligne_temps + 2, colonne_opération, colonne_temps, colonne_opération, colonne_vitesse, colonne_bt5)
        Objet_commentaire_dosage = CoordonnéesExcelGamme(wb, 'Total', ligne_1 + 1, colonne_produit, colonne_produit, colonne_anc_code, colonne_poids)


        Liste_mode_operatoire_1temps = Objet_mode_operatoire_1temps.get_tout()
        Liste_composants_code_poids = Objet_composants.get_tout()
        Liste_commentaire_dosage = Objet_commentaire_dosage.get_tout()
        Liste_ancien_code = Objet_ancien_code.get_tout()

        if temps2:
            Objet_ancien_code_2temps = CoordonnéesExcel(wb, 'Mélange complet', ligne_temps2 + 1, colonne_produit, colonne_produit, colonne_anc_code)
            Objet_commentaire_dosage_2temps = CoordonnéesExcelGamme(wb, 'Mélange complet', ligne_temps2 + 1,
                                                                    colonne_produit,
                                                                    colonne_produit, colonne_anc_code, colonne_poids)
            Objet_mode_operatoire_2temps = CoordonnéesExcelGamme2Temps(wb, 'Densité', ligne_temps2 + 3, colonne_opération, colonne_temps, colonne_opération, (colonne_opération, 1), colonne_vitesse, colonne_bt5)
            Liste_ancien_code_2temps = Objet_ancien_code_2temps.get_tout()
            Liste_commentaire_dosage_2temps = Objet_commentaire_dosage_2temps.get_tout()
            Liste_mode_operatoire_2temps = Objet_mode_operatoire_2temps.get_tout()
            Liste_composants_code_poids_2temps = Objet_composants_2temps.get_tout()
            Liste_ancien_code += Liste_ancien_code_2temps
            Liste_composants_code_poids += Liste_composants_code_poids_2temps
            Liste_commentaire_dosage += Liste_commentaire_dosage_2temps

        ########################################################################################################################
        # Les tables sont remises en forme afin d'etre injecter dans les commentaires des opérations
        ########################################################################################################################
        try:
            tableOP2 = ""
            if temps2:
                tableOP2 = tableize2(pd.DataFrame(Liste_mode_operatoire_2temps,
                                                 columns=Objet_mode_operatoire_2temps.get_colonnes()))
            tableOp = tableize(pd.DataFrame(Liste_mode_operatoire_1temps, columns=["Temps (s)", "Operation", "Vitesse(RPM)", "BT5( C)"]))

            # tableComp = tableize(pd.DataFrame(listeComp, columns=[colonne_name, "Poids"]))
            tableDosage = tableize(pd.DataFrame(Liste_commentaire_dosage, columns=["Désignation", "Code", "Poids( Kg)"]))
        except Exception as e:
            log_exceptions.append((book, e))
            continue

        ########################################################################################################################
        for poids_code in Liste_ancien_code:
            if str(poids_code[1]) in table_anc_code:
                pass
            else:
                table_anc_code[str(poids_code[1])] = str(poids_code[0])

        ########################################################################################################################
        try: # Si le code du composant n'est pas un code Sylob, on le converti en se basant sur la liste de correspondance
            for code_poids in Liste_composants_code_poids:
                if code_poids[0].startswith("A"):
                    continue
                code_poids[0] = liste_composants_Sylob[code_poids[0]]

        except Exception:
            log_sans_code_Sylob.append((book, f'{code_poids[0]} n\'a pas de code Sylob'))
            continue
        ########################################################################################################################
        # Convertion des temps en cadence
        ########################################################################################################################
        if temps_OP == '0':
            pass
        else:
            temps_OP = str(((float(poids_total)*1000)/float(temps_OP))*3.6)
        if temps_OP_2 == '0':
            pass
        else:
            temps_OP_2 = str(((float(poids_total)*1000)/float(temps_OP_2))*3.6)
        ########################################################################################################################
        # Extraction des centres de charge
        centre_de_charge = str(wb.iloc[int(ligne_1) - 1 + 1][ord(colonne_temps) - 97 + 2])
        try:
            if centre_de_charge == 'nan':
                centre_de_charge = str(wb.iloc[int(ligne_1) - 1 + 2][ord(colonne_temps) - 97 + 2])
                if centre_de_charge == 'nan':
                    raise Exception
                else:
                    centre_de_charge = centre_de_charge_acceleration
            else:
                centre_de_charge = centre_de_charge_melangeage
        except Exception:
            print((book, 'pas de centre de charge'))
            continue
        yes_ = True
        ########################################################################################################################
        # À ce niveau, on se base sur les dictionnaires EDI afin d'editer le contenu en fonction des champs. L'utilisation des dictionnaires nous permet d'acceder facilement aux clés de modifier le contenu et de grader l'ordre définit d'informations
        # On converti les dictionnaires en listes contenants uniquement les valeurs du dictionnaire
        ########################################################################################################################


        try:# On renseigne le code article
            EDI_article_dict["Code_Art"] = table_codes_articles[book1]
        except Exception:
            # log_exceptions.append((book, "Pas de code Sylob Mélange"))
            pass
        EDI_article_dict["Code_Art"] = "A03-00385" # Pour le test, normalement c'est repris dans les lignes juste avant
        EDI_composants_dict["Compo_Qté_Pour"] = poids_total # Renseigne le poids du batch
        EDI_article_dict["CODE_DON_TECH"] = codeDT + book1 # // code, libellée de la DT
        EDI_article_dict["LIB"] = libDT + book1
        EDI_article_dict["Ind_Art"] = "" # Indice de l'article
        EDI_article_liste = [EDI_article_dict[key] for key in EDI_article_dict] # Dict --> List

        EDI_opération_dict["Opé_Lib"] = "DOSAGE"  # On change le nom de l'opération
        EDI_opération_dict["Opé_Centre"] = centre_de_charge_dosage  # On change le centre de charge
        EDI_opération_dict["Temps_Fab"] = '0'
        EDI_opération_dict["Opé_Comm"] = tableDosage # Mise en commentaire du dosage. Le nom de l'operation est par défaut Dosage
        EDI_opération_liste.append([EDI_opération_dict[key] for key in EDI_opération_dict]) # Dict --> List
        EDI_opération_liste[0] = EDI_article_liste + EDI_opération_liste[0] # Les lignes de l'EDI sont structurées en "Infos article" + "Info de DT"


        EDI_opération_dict["Opé_Lib"] = "MELANGEAGE" # On change le nom de l'opération
        EDI_opération_dict["Opé_Centre"] = centre_de_charge  # On change le centre de charge
        EDI_opération_dict["Temps_Fab"] = temps_OP
        EDI_opération_dict["Opé_Comm"] = tableOp # Mise en commentaire du mélangeage
        EDI_opération_liste.append([EDI_opération_dict[key] for key in EDI_opération_dict]) # Dict --> List
        EDI_opération_liste[1] = EDI_article_liste + EDI_opération_liste[1] # Les lignes de l'EDI sont structurées en "Infos article" + "Info de DT"

        if temps2:
            EDI_opération_dict["Opé_Centre"] = centre_de_charge_acceleration  # On change le centre de charge
        else:
            EDI_opération_dict["Opé_Centre"] = centre_de_charge_melangeage  # On change le centre de charge
        EDI_opération_dict["Opé_Lib"] = "ACCELERATION" # On change le nom de l'opération
        EDI_opération_dict["Temps_Fab"] = temps_OP_2
        EDI_opération_dict["Opé_Comm"] = tableOP2
        # EDI_opération_dict["Opé_Comm"] = tableOp # Mise en commentaire du mélangeage
        EDI_opération_liste.append([EDI_opération_dict[key] for key in EDI_opération_dict]) # Dict --> List
        EDI_opération_liste[2] = EDI_article_liste + EDI_opération_liste[2] # Les lignes de l'EDI sont structurées en "Infos article" + "Info de DT"

        i = 0
        for code_poids in Liste_composants_code_poids:
            EDI_composants_dict["Compo_Code"] = code_poids[0] # On renseigne le code du composant
            EDI_composants_dict["Compo_Qté"] = code_poids[1] # La quantité
            EDI_composants_liste.append([EDI_composants_dict[key] for key in EDI_composants_dict])#  Dict --> List
            EDI_composants_liste[i] = EDI_article_liste + EDI_composants_liste[i] # Les lignes de l'EDI sont structurées en "Infos article" + "Info de DT"
            i += 1

        liste_csv = liste_csv + EDI_opération_liste + EDI_composants_liste # On rajoute les informations à la liste csv globale
        liste_blanche.append(book)

        # for name in vars().keys():
        #     print(name)
    # print(dictio_centre1)
    # print(dictio_centre2)
    # print('ouvert')
    # print(dictio_ouvert)
    liste_anc_code = []
    liste_anc_code = [(key, table_anc_code[key]) for key in table_anc_code]
    # print(liste_anc_code)
    return liste_csv, liste_noire, liste_blanche, liste_vari, log_exceptions, table_anc_code, log_sans_code_Sylob#liste_anc_code
