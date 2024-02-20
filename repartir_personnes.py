import openpyxl
import random

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
