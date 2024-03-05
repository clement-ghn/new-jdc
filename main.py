# main.py
import tkinter as tk
import openpyxl
from repartir_personnes import repartir_personnes
from decouper_groupes import decouper_groupes


wb = openpyxl.load_workbook("volontaires.xlsx")
sheet = wb.active

root = tk.Tk()
root.title("Répartition des Groupes")

nombre_groupes = 2


# Fonction pour effectuer le tirage et afficher le résultat
def tirage_et_affichage():
    nombre_sous_groupes_a = int(entry_a.get())
    nombre_sous_groupes_b = int(entry_b.get())
    
    # Capacités des maisons
    capacite_homme_input = entry_garcon.get()
    capacite_femme_input = entry_fille.get()
    
    capacite_maison_homme = list(map(int, capacite_homme_input.split(',')))
    capacite_maison_femme = list(map(int, capacite_femme_input.split(',')))
    
    # Maisons pour chaque sexe
    maisons_homme = [[] for _ in range(len(capacite_maison_homme))]
    maisons_femme = [[] for _ in range(len(capacite_maison_femme))]
    
    groupes = repartir_personnes(sheet, nombre_groupes)
    sous_groupes = decouper_groupes(groupes, nombre_sous_groupes_a, nombre_sous_groupes_b)

    bouton_tirage.destroy()
    label_a.destroy()
    entry_a.destroy()
    label_b.destroy()
    entry_b.destroy()
    label_garcon.destroy()
    entry_garcon.destroy()
    label_fille.destroy()
    entry_fille.destroy()

    # Create a canvas in the root window
    canvas = tk.Canvas(root)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Add a scrollbar for the canvas
    scrollbar = tk.Scrollbar(root, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.configure(yscrollcommand=scrollbar.set)

    # Function to resize the canvas scrolling region
    def resize_canvas(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    # Bind the resize_canvas function to the <Configure> event of the canvas
    canvas.bind("<Configure>", resize_canvas)

    # Create a frame inside the canvas to contain all the group frames
    container = tk.Frame(canvas)
    container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # Add the container frame to the canvas
    canvas.create_window((0, 0), window=container, anchor="nw")

    # Now, create and add all the group frames to the container frame
    for i, groupe in enumerate(sous_groupes):
        group_frame = tk.Frame(container)
        group_frame.grid(row=0, column=i, padx=10, pady=10, sticky="nw")

        section_label = tk.Label(group_frame, text=f"Compagnie {'Alpha' if i == 0 else 'Bravo'}")
        section_label.grid(row=0, column=0, padx=10, pady=5, sticky="nw")

        for j, sous_groupe in enumerate(groupe):
            subgroup_label = tk.Label(group_frame, text=f"{'A' if i == 0 else 'B'}{j+1}")
            subgroup_label.grid(row=j*2+1, column=0, padx=10, pady=5, sticky="nw")  # Utilisation de j*2+1 pour espacer les libellés

            subgroup_frame = tk.Frame(group_frame)
            subgroup_frame.grid(row=j*2+2, column=0, padx=10, pady=5, sticky="nw")

            sublistbox = tk.Listbox(subgroup_frame, width=30, height=15)
            sublistbox.pack(side=tk.LEFT, fill=tk.BOTH)
            
            for personne in sous_groupe:
                sublistbox.insert(tk.END, f"{personne['prenom']} {personne['nom']}, {personne['date_naissance']}, {personne['sexe']}, {personne['email']}, {personne['telephone']}, {personne['code_postal']}, {personne['ville']}, {personne['departement']}, {personne['niveau']}, {personne['bus_no']}")
                if personne['sexe'] == 'Homme':
                    for k, capacite in enumerate(capacite_maison_homme):
                        if capacite > 0:
                            maisons_homme[k].append([personne['prenom'], personne['nom']])  # Ajoutez une liste de données de personne
                            capacite_maison_homme[k] -= 1
                            break
                elif personne['sexe'] == 'Femme':
                    for k, capacite in enumerate(capacite_maison_femme):
                        if capacite > 0:
                            maisons_femme[k].append([personne['prenom'], personne['nom']])  # Ajoutez une liste de données de personne
                            capacite_maison_femme[k] -= 1
                            break

    # Ajout des widgets pour afficher les maisons
    for k, maison in enumerate(maisons_homme):
        maison_frame = tk.Frame(container)
        maison_frame.grid(row=2*k+2, column=0, padx=10, pady=5, sticky="nw")
        label_maison_homme = tk.Label(maison_frame, text=f"Maison Homme {k+1}")
        label_maison_homme.pack(side=tk.TOP)
        sublistbox_maison_homme = tk.Listbox(maison_frame, width=30, height=5)
        sublistbox_maison_homme.pack(side=tk.LEFT, fill=tk.BOTH)
        for personne in maison:
            sublistbox_maison_homme.insert(tk.END, f"{personne[0]} {personne[1]}")
        scrollbar_maison_homme = tk.Scrollbar(maison_frame, orient=tk.VERTICAL, command=sublistbox_maison_homme.yview)
        scrollbar_maison_homme.pack(side=tk.RIGHT, fill=tk.Y)
        sublistbox_maison_homme.configure(yscrollcommand=scrollbar_maison_homme.set)
    
    for k, maison in enumerate(maisons_femme):
        maison_frame = tk.Frame(container)
        maison_frame.grid(row=2*k+2, column=1, padx=10, pady=5, sticky="nw")
        label_maison_femme = tk.Label(maison_frame, text=f"Maison Femme {k+1}")
        label_maison_femme.pack(side=tk.TOP)
        sublistbox_maison_femme = tk.Listbox(maison_frame, width=30, height=5)
        sublistbox_maison_femme.pack(side=tk.LEFT, fill=tk.BOTH)
        for personne in maison:
            sublistbox_maison_femme.insert(tk.END, f"{personne[0]} {personne[1]}")
        scrollbar_maison_femme = tk.Scrollbar(maison_frame, orient=tk.VERTICAL, command=sublistbox_maison_femme.yview)
        scrollbar_maison_femme.pack(side=tk.RIGHT, fill=tk.Y)
        sublistbox_maison_femme.configure(yscrollcommand=scrollbar_maison_femme.set)

    # Update the canvas scrolling region
    resize_canvas(None)


#Tkinter part
label_a = tk.Label(root, text="Nombre de sous-groupes pour Groupe A:")
label_a.grid(row=0, column=0, padx=10, pady=5, sticky="w")
entry_a = tk.Entry(root)
entry_a.grid(row=0, column=1, padx=10, pady=5)

label_b = tk.Label(root, text="Nombre de sous-groupes pour Groupe B:")
label_b.grid(row=1, column=0, padx=10, pady=5, sticky="w")
entry_b = tk.Entry(root)
entry_b.grid(row=1, column=1, padx=10, pady=5)

label_garcon = tk.Label(root, text="Capacités maisons garçons:")
label_garcon.grid(row=0, column=2, padx=10, pady=5, sticky="w")
entry_garcon = tk.Entry(root)
entry_garcon.grid(row=0, column=3, padx=10, pady=5)

label_fille = tk.Label(root, text="Capacités maisons filles:")
label_fille.grid(row=1, column=2, padx=10, pady=5, sticky="w")
entry_fille = tk.Entry(root)
entry_fille.grid(row=1, column=3, padx=10, pady=5)


bouton_tirage = tk.Button(root, text="Effectuer le tirage", command=tirage_et_affichage)
bouton_tirage.grid(row=2, columnspan=2, pady=10)

root.mainloop()
