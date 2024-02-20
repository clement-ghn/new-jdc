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