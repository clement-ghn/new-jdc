from math import ceil,floor

def decouper_groupes(groupes, nombre_sous_groupes_a, nombre_sous_groupes_b):
    sous_groupes = [[] for _ in range(len(groupes))]
    for i, groupe in enumerate(groupes):
        hommes = [personne for personne in groupe if personne['sexe'] == 'Homme']
        femmes = [personne for personne in groupe if personne['sexe'] == 'Femme']

        taille_sous_groupe_a = ceil(len(hommes) / nombre_sous_groupes_a)
        taille_sous_groupe_b = ceil(len(femmes) / nombre_sous_groupes_b)




        hommes_restants = hommes.copy()
        femmes_restantes = femmes.copy()

        if i == 0:
            # Groupe A
            for j in range(nombre_sous_groupes_a):
                nb_hommes = min(taille_sous_groupe_a, len(hommes_restants))
                nb_femmes = min(taille_sous_groupe_b, len(femmes_restantes))
                sous_groupe_h = hommes_restants[:nb_hommes]
                sous_groupe_f = femmes_restantes[:nb_femmes]
                sous_groupe = sous_groupe_h + sous_groupe_f
                sous_groupes[i].append(sous_groupe)
                hommes_restants = hommes_restants[nb_hommes:]
                femmes_restantes = femmes_restantes[nb_femmes:]
        else:
            # Groupe B
            for j in range(nombre_sous_groupes_b):
                nb_hommes = min(taille_sous_groupe_a, len(hommes_restants))
                nb_femmes = min(taille_sous_groupe_b, len(femmes_restantes))
                sous_groupe_h = hommes_restants[:nb_hommes]
                sous_groupe_f = femmes_restantes[:nb_femmes]
                sous_groupe = sous_groupe_h + sous_groupe_f
                sous_groupes[i].append(sous_groupe)
                hommes_restants = hommes_restants[nb_hommes:]
                femmes_restantes = femmes_restantes[nb_femmes:]

    return sous_groupes