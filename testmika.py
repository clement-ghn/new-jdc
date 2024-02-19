import openpyxl
import random
import tkinter as tk

wb = openpyxl.load_workbook("volontaires.xlsx")
sheet = wb.active

root = tk.Tk()
root.title("Répartition des Groupes")

nombre_groupes = 2
nombre_sous_groupes = 4

# Fonction pour répartir aléatoirement les personnes en groupes
def repartir_personnes(sheet, nombre_groupes):
    groupes = [[] for _ in range(nombre_groupes)]
    personnes = []
    for row in sheet.iter_rows(values_only=True):
        _id, cohorte, prenom, nom, date_naissance, ville_naissance, sexe, email, telephone, adresse_postale, code_postal, ville, departement, region, academie, pays, nom_hebergeur, prenom_hebergeur, lien_hebergeur, adresse_etranger, code_postal_etranger, ville_etranger, pays_etranger, situation, niveau, type_etablissement, nom_etablissement, code_postal_etablissement, ville_etablissement, departement_etablissement, uai_etablissement, quartier_prioritaire_ville, zone_rurale, handicap, beneficiaire_pps, beneficiaire_pai, structure_medico_sociale, nom_structure_medico_sociale, adresse_structure_medico_sociale, code_postal_structure_medico_sociale, ville_structure_medico_sociale, amenagement_specifique, nature_amenagement_specifique, amenagement_mobilite_reduite, besoin_affectation_departement_residence, allergies_intolerances_alimentaires, activite_haut_niveau, nature_activite_haut_niveau, activites_haut_niveau_affectation_departement_residence, document_activite_haut_niveau, consentement_representants_legaux, droit_image_accord, droit_image_statut, autotest_pcr_accord, autotest_pcr_statut, reglement_interieur, fiche_sanitaire, presence_arrivee, date_depart, motif_depart, commentaire_depart, statut_representant_legal_1, prenom_representant_legal_1, nom_representant_legal_1, email_representant_legal_1, telephone_representant_legal_1, adresse_representant_legal_1, code_postal_representant_legal_1, ville_representant_legal_1, departement_representant_legal_1, region_representant_legal_1, statut_representant_legal_2, prenom_representant_legal_2, nom_representant_legal_2, email_representant_legal_2, telephone_representant_legal_2, adresse_representant_legal_2, code_postal_representant_legal_2, ville_representant_legal_2, departement_representant_legal_2, region_representant_legal_2, derniere_connexion, statut_general, statut_phase_1, raison_desistement, message_desistement, id_centre, code_centre_2021, code_centre_2022, nom_centre, ville_centre, departement_centre, region_centre, participation_sejour, confirmation_point_rassemblement, se_rend_centre_propres_moyens, informations_transport_transmises_services_locaux, bus_no, adresse_point_rassemblement, date_aller, date_retour = row
        personne = {
            '_id': _id,
            'cohorte': cohorte,
            'prenom': prenom,
            'nom': nom,
            'date_naissance': date_naissance,
            'ville_naissance': ville_naissance,
            'sexe': sexe,
            'email': email,
            'telephone': telephone,
            'adresse_postale': adresse_postale,
            'code_postal': code_postal,
            'ville': ville,
            'departement': departement,
            'region': region,
            'academie': academie,
            'pays': pays,
            'nom_hebergeur': nom_hebergeur,
            'prenom_hebergeur': prenom_hebergeur,
            'lien_hebergeur': lien_hebergeur,
            'adresse_etranger': adresse_etranger,
            'code_postal_etranger': code_postal_etranger,
            'ville_etranger': ville_etranger,
            'pays_etranger': pays_etranger,
            'situation': situation,
            'niveau': niveau,
            'type_etablissement': type_etablissement,
            'nom_etablissement': nom_etablissement,
            'code_postal_etablissement': code_postal_etablissement,
            'ville_etablissement': ville_etablissement,
            'departement_etablissement': departement_etablissement,
            'uai_etablissement': uai_etablissement,
            'quartier_prioritaire_ville': quartier_prioritaire_ville,
            'zone_rurale': zone_rurale,
            'handicap': handicap,
            'beneficiaire_pps': beneficiaire_pps,
            'beneficiaire_pai': beneficiaire_pai,
            'structure_medico_sociale': structure_medico_sociale,
            'nom_structure_medico_sociale': nom_structure_medico_sociale,
            'adresse_structure_medico_sociale': adresse_structure_medico_sociale,
            'code_postal_structure_medico_sociale': code_postal_structure_medico_sociale,
            'ville_structure_medico_sociale': ville_structure_medico_sociale,
            'amenagement_specifique': amenagement_specifique,
            'nature_amenagement_specifique': nature_amenagement_specifique,
            'amenagement_mobilite_reduite': amenagement_mobilite_reduite,
            'besoin_affectation_departement_residence': besoin_affectation_departement_residence,
            'allergies_intolerances_alimentaires': allergies_intolerances_alimentaires,
            'activite_haut_niveau': activite_haut_niveau,
            'nature_activite_haut_niveau': nature_activite_haut_niveau,
            'activites_haut_niveau_affectation_departement_residence': activites_haut_niveau_affectation_departement_residence,
            'document_activite_haut_niveau': document_activite_haut_niveau,
            'consentement_representants_legaux': consentement_representants_legaux,
            'droit_image_accord': droit_image_accord,
            'droit_image_statut': droit_image_statut,
            'autotest_pcr_accord': autotest_pcr_accord,
            'autotest_pcr_statut': autotest_pcr_statut,
            'reglement_interieur': reglement_interieur,
            'fiche_sanitaire': fiche_sanitaire,
            'presence_arrivee': presence_arrivee,
            'date_depart': date_depart,
            'motif_depart': motif_depart,
            'commentaire_depart': commentaire_depart,
            'statut_representant_legal_1': statut_representant_legal_1,
            'prenom_representant_legal_1': prenom_representant_legal_1,
            'nom_representant_legal_1': nom_representant_legal_1,
            'email_representant_legal_1': email_representant_legal_1,
            'telephone_representant_legal_1': telephone_representant_legal_1,
            'adresse_representant_legal_1': adresse_representant_legal_1,
            'code_postal_representant_legal_1': code_postal_representant_legal_1,
            'ville_representant_legal_1': ville_representant_legal_1,
            'departement_representant_legal_1': departement_representant_legal_1,
            'region_representant_legal_1': region_representant_legal_1,
            'statut_representant_legal_2': statut_representant_legal_2,
            'prenom_representant_legal_2': prenom_representant_legal_2,
            'nom_representant_legal_2': nom_representant_legal_2,
            'email_representant_legal_2': email_representant_legal_2,
            'telephone_representant_legal_2': telephone_representant_legal_2,
            'adresse_representant_legal_2': adresse_representant_legal_2,
            'code_postal_representant_legal_2': code_postal_representant_legal_2,
            'ville_representant_legal_2': ville_representant_legal_2,
            'departement_representant_legal_2': departement_representant_legal_2,
            'region_representant_legal_2': region_representant_legal_2,
            'derniere_connexion': derniere_connexion,
            'statut_general': statut_general,
            'statut_phase_1': statut_phase_1,
            'raison_desistement': raison_desistement,
            'message_desistement': message_desistement,
            'id_centre': id_centre,
            'code_centre_2021': code_centre_2021,
            'code_centre_2022': code_centre_2022,
            'nom_centre': nom_centre,
            'ville_centre': ville_centre,
            'departement_centre': departement_centre,
            'region_centre': region_centre,
            'participation_sejour': participation_sejour,
            'confirmation_point_rassemblement': confirmation_point_rassemblement,
            'se_rend_centre_propres_moyens': se_rend_centre_propres_moyens,
            'informations_transport_transmises_services_locaux': informations_transport_transmises_services_locaux,
            'bus_no': bus_no,
            'adresse_point_rassemblement': adresse_point_rassemblement,
            'date_aller': date_aller,
            'date_retour': date_retour
        }
        personnes.append(personne)

    random.shuffle(personnes)  # Mélanger aléatoirement les personnes

    for i, personne in enumerate(personnes):
        groupe_index = i % nombre_groupes
        groupes[groupe_index].append(personne)

    return groupes



# Fonction pour découper chaque groupe en sous-groupes
def decouper_groupes(groupes, nombre_sous_groupes):
    sous_groupes = [[] for _ in range(len(groupes))]
    for i, groupe in enumerate(groupes):
        taille_groupe = len(groupe)
        taille_sous_groupe = taille_groupe // nombre_sous_groupes
        for j in range(nombre_sous_groupes):
            debut = j * taille_sous_groupe
            fin = (j + 1) * taille_sous_groupe if j < nombre_sous_groupes - 1 else taille_groupe
            sous_groupe = groupe[debut:fin]
            sous_groupes[i].append(sous_groupe)
    return sous_groupes


def chambres_maisons(sous_groupes, nombre_chambres):
    chambres_2 = nombre_chambres // 2
    chambres_3 = nombre_chambres - chambres_2

    for groupe in sous_groupes:
        for i in range(chambres_2):
            groupe[2 * i].append("chambre 2")
        for i in range(chambres_3):
            groupe[2 * i + 1].append("chambre 3")

    return sous_groupes





# Fonction pour effectuer le tirage et afficher le résultat
def tirage_et_affichage():
    #nb_chambre_2 = int(select_chambre2.get())
    #nb_chambre_3 = int(select_chambre3.get())
    #nombre_chambres = nb_chambre_2 + nb_chambre_3

    # Exemple d'utilisation
    groupes = repartir_personnes(sheet, nombre_groupes)
    sous_groupes = decouper_groupes(groupes, nombre_sous_groupes)
    #sous_groupes = chambres_maisons(sous_groupes, nombre_chambres)

    noms_maisons = ["anemones", "brimbelles", "genets", "jonquilles", "bruyeres", "sapins", "soyotte", "rien"]
    

    for i, groupe in enumerate(sous_groupes):
        group_frame = tk.Frame(root)
        group_frame.grid(row=0, column=i, padx=10, pady=10, sticky="nw")

        section_label = tk.Label(group_frame, text=f"Groupe {'A' if i == 0 else 'B'}")
        section_label.grid(row=0, column=0, padx=10, pady=5, sticky="nw")

        for j, sous_groupe in enumerate(groupe):
            subgroup_label = tk.Label(group_frame, text= noms_maisons[j] if i ==0 else noms_maisons[j+4])
            subgroup_label.grid(row=j*2+1, column=0, padx=10, pady=5, sticky="nw")  # Utilisation de j*2+1 pour espacer les libellés

            subgroup_frame = tk.Frame(group_frame)
            subgroup_frame.grid(row=j*2+2, column=0, padx=10, pady=5, sticky="nw")
            scrollbar = tk.Scrollbar(subgroup_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            sublistbox = tk.Listbox(subgroup_frame, yscrollcommand=scrollbar.set, width=30, height=10)
            sublistbox.pack(side=tk.LEFT, fill=tk.BOTH)
            scrollbar.config(command=sublistbox.yview)
            for personne in sous_groupe:
                sublistbox.insert(tk.END, f"{personne['prenom']} {personne['nom']}, {personne['sexe']}, {personne['bus_no']}")




#label_chambre2 = tk.Label(root, text="Nombre de chambres de 2")
#label_chambre2.grid(row=0, column=0)
    
#select_chambre2 = tk.Spinbox(root, from_=0, to=100)
#select_chambre2.grid(row=0, column=1)

#abel_chambre3 = tk.Label(root, text="Nombre de chambres de 3")
#label_chambre3.grid(row=0, column=3)

#select_chambre3 = tk.Spinbox(root, from_=0, to=100)
#select_chambre3.grid(row=0, column=4)



bouton_tirage = tk.Button(root, text="Effectuer le tirage", command=tirage_et_affichage)
bouton_tirage.grid(row=1, columnspan=2)

root.mainloop()
