import os
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog

# Fonction pour afficher les e-mails
def afficher_emails():
    destinataires = destinataires_entry.get().split(';')
    copie = copie_entry.get().split(';')
    corps_mail = corps_entry.get("1.0", "end-1c")
    repertoire_pj = repertoire_entry.get()

    # Définir la signature
    signature = "\n\nCordialement,\nVotre Nom\nVotre Poste\nVotre Entreprise\nVotre Téléphone\nVotre Email"

    # Ajouter la signature au corps du mail
    corps_mail += signature

    outlook = win32.Dispatch('outlook.application')

    for nom_fichier in os.listdir(repertoire_pj):
        if nom_fichier.endswith(('.pdf', '.docx', '.xlsx')):
            sujet = f"DEMANDE RÉGLEMENT FACTURE - {nom_fichier}"
            chemin_fichier = os.path.join(repertoire_pj, nom_fichier)

            mail = outlook.CreateItem(0)
            mail.Subject = sujet
            mail.Body = corps_mail
            mail.To = '; '.join(destinataires)
            mail.CC = '; '.join(copie)
            mail.Attachments.Add(chemin_fichier)

            # Afficher l'e-mail au lieu de l'envoyer
            mail.Display()

    messagebox.showinfo("Information", "Les e-mails ont été affichés dans Outlook.")

# Fonction pour parcourir et sélectionner un répertoire
def parcourir_repertoire():
    repertoire = filedialog.askdirectory()
    if repertoire:
        repertoire_entry.delete(0, tk.END)  # Effacer l'entrée actuelle
        repertoire_entry.insert(0, repertoire)  # Insérer le nouveau chemin

# Création de la fenêtre principale
app = tk.Tk()
app.title("Affichage d'e-mails avec Outlook")
app.geometry("500x550")  # Hauteur augmentée à 550
app.configure(bg="#f0f0f0")

# Titre
title_label = tk.Label(app, text="Affichage d'e-mails avec Outlook", font=("Arial", 16, "bold"), bg="#f0f0f0")
title_label.pack(pady=10)

# Champs de saisie
tk.Label(app, text="Destinataires (séparés par ;) :", bg="#f0f0f0").pack(pady=5)
destinataires_entry = tk.Entry(app, width=50)
destinataires_entry.pack(pady=5)

tk.Label(app, text="Copie (séparés par ;) :", bg="#f0f0f0").pack(pady=5)
copie_entry = tk.Entry(app, width=50)
copie_entry.pack(pady=5)

tk.Label(app, text="Corps du mail :", bg="#f0f0f0").pack(pady=5)
corps_entry = scrolledtext.ScrolledText(app, width=50, height=10)
corps_entry.pack(pady=5)

tk.Label(app, text="Répertoire des pièces jointes :", bg="#f0f0f0").pack(pady=5)
repertoire_entry = tk.Entry(app, width=50)
repertoire_entry.pack(pady=5)

# Bouton pour parcourir le répertoire
parcourir_button = tk.Button(app, text="Parcourir", command=parcourir_repertoire, bg="#2196F3", fg="white")
parcourir_button.pack(pady=5)

# Bouton pour afficher les e-mails
afficher_button = tk.Button(app, text="Afficher les e-mails", command=afficher_emails, bg="#4CAF50", fg="white", font=("Arial", 12))
afficher_button.pack(pady=20)

# Lancer l'application
app.mainloop()
