import openpyxl
import random
import tkinter as tk

wb = openpyxl.load_workbook("volontaires.xlsx")
sheet = wb.active

root = tk.Tk()
root.title("Répartition des Groupes")

nombre_groupes = 2
#nombre_sous_groupes_a = 4
#nombre_sous_groupes_b = 4


# Fonction pour répartir aléatoirement les personnes en groupes
def repartir_personnes(sheet, nombre_groupes):
    groupes = [[] for _ in range(nombre_groupes)]
    personnes = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
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
            'pays': pays,
            'fiche_sanitaire': fiche_sanitaire,
            'presence_arrivee': presence_arrivee,
            'date_depart': date_depart,
            'motif_depart': motif_depart,
            'commentaire_depart': commentaire_depart,
            'statut_general': statut_general,
            'id_centre': id_centre,
            'code_centre_2021': code_centre_2021,
            'code_centre_2022': code_centre_2022,
            'nom_centre': nom_centre,
            'ville_centre': ville_centre,
            'departement_centre': departement_centre,
            'region_centre': region_centre,
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



def decouper_groupes(groupes, nombre_sous_groupes_a, nombre_sous_groupes_b):
    sous_groupes = [[] for _ in range(len(groupes))]
    for i, groupe in enumerate(groupes):
        taille_groupe = len(groupe)
        taille_sous_groupe_a = taille_groupe // nombre_sous_groupes_a
        taille_sous_groupe_b = taille_groupe // nombre_sous_groupes_b

        if i == 0:
            # Groupe A
            for j in range(nombre_sous_groupes_a):
                debut = j * taille_sous_groupe_a
                fin = (j + 1) * taille_sous_groupe_a if j < nombre_sous_groupes_a - 1 else taille_groupe
                sous_groupe = groupe[debut:fin]
                sous_groupes[i].append(sous_groupe)
        else:
            # Groupe B
            for j in range(nombre_sous_groupes_b):
                debut = j * taille_sous_groupe_b
                fin = (j + 1) * taille_sous_groupe_b if j < nombre_sous_groupes_b - 1 else taille_groupe
                sous_groupe = groupe[debut:fin]
                sous_groupes[i].append(sous_groupe)

    return sous_groupes

# Fonction pour effectuer le tirage et afficher le résultat
def tirage_et_affichage():
    nombre_sous_groupes_a = int(entry_a.get())
    nombre_sous_groupes_b = int(entry_b.get())
    groupes = repartir_personnes(sheet, nombre_groupes)
    sous_groupes = decouper_groupes(groupes, nombre_sous_groupes_a, nombre_sous_groupes_b)
    bouton_tirage.destroy()
    label_a.destroy()
    entry_a.destroy()
    label_b.destroy()
    entry_b.destroy()
    #noms_maisons = ["Anemones", "Brimbelles", "Genets", "Jonquilles", "Bruyeres", "Sapins", "Soyotte", "tableau en plus"]
    noms_maisons = [maison.strip() for maison in entry_c.get().split(",")]
    label_c.destroy()
    entry_c.destroy()

    for i, groupe in enumerate(sous_groupes):
        group_frame = tk.Frame(root)
        group_frame.grid(row=0, column=i, padx=10, pady=10, sticky="nw")

        section_label = tk.Label(group_frame, text=f"Groupe {'A' if i == 0 else 'B'}")
        section_label.grid(row=0, column=0, padx=10, pady=5, sticky="nw")

        for j, sous_groupe in enumerate(groupe):
            subgroup_label = tk.Label(group_frame, text= noms_maisons[j] if i == 0 else noms_maisons[j+nombre_sous_groupes_a])
            subgroup_label.grid(row=j*2+1, column=0, padx=10, pady=5, sticky="nw")  # Utilisation de j*2+1 pour espacer les libellés

            subgroup_frame = tk.Frame(group_frame)
            subgroup_frame.grid(row=j*2+2, column=0, padx=10, pady=5, sticky="nw")
            scrollbar = tk.Scrollbar(subgroup_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            sublistbox = tk.Listbox(subgroup_frame, yscrollcommand=scrollbar.set, width=30, height=5)
            sublistbox.pack(side=tk.LEFT, fill=tk.BOTH)
            scrollbar.config(command=sublistbox.yview)
            for personne in sous_groupe:
                sublistbox.insert(tk.END, f"{personne['prenom']} {personne['nom']}, {personne['sexe']}, {personne['bus_no']}")


#Tkinter part

label_a = tk.Label(root, text="Nombre de sous-groupes pour Groupe A:")
label_a.grid(row=0, column=0, padx=10, pady=5, sticky="w")
entry_a = tk.Entry(root)
entry_a.grid(row=0, column=1, padx=10, pady=5)

label_b = tk.Label(root, text="Nombre de sous-groupes pour Groupe B:")
label_b.grid(row=1, column=0, padx=10, pady=5, sticky="w")
entry_b = tk.Entry(root)
entry_b.grid(row=1, column=1, padx=10, pady=5)

label_c = tk.Label(root, text="Nom des maisons du groupe A")
label_c.grid(row=0, column=2, padx=10, pady=5, sticky="w")
entry_c = tk.Entry(root)
entry_c.grid(row=0, column=3, padx=10, pady=5)


bouton_tirage = tk.Button(root, text="Effectuer le tirage", command=tirage_et_affichage)
bouton_tirage.grid(row=2, columnspan=2, pady=10)

root.mainloop()