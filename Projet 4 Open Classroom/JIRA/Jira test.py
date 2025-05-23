import win32com.client
import pandas as pd
import re

# Fonction pour se connecter à un sous-dossier spécifique dans Outlook
def get_outlook_subfolder(email_address, parent_folder_name, target_folder_name):
    try:
        # Connexion à Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        # Accéder au dossier parent (par exemple, la Boîte de réception)
        root_folder = outlook.Folders(email_address).Folders(parent_folder_name)
        # Lister tous les sous-dossiers pour s'assurer que le dossier existe
        print(f"Dossiers sous {parent_folder_name} dans {email_address} :")
        for folder in root_folder.Folders:
            print("-", folder.Name)
        # Accéder au dossier spécifique
        target_folder = root_folder.Folders(target_folder_name)
        return target_folder
    except Exception as e:
        print(f"Erreur lors de l'accès au dossier {target_folder_name}: {e}")
        return None

# Adresse email, dossier parent, et dossier cible
email_address = "testjiraaaaaaaaaa@outlook.fr"
parent_folder_name = "Boîte de réception"  # Adapter si nécessaire
target_folder_name = "JIRA"  # Nom du sous-dossier contenant les mails JIRA

# Connexion au sous-dossier Outlook
folder = get_outlook_subfolder(email_address, parent_folder_name, target_folder_name)

# Si le dossier est trouvé, procéder à l'extraction
if folder:
    # Initialiser une liste pour stocker les données extraites
    tickets_data = []

    # Parcourir les mails dans le dossier
    for mail in folder.Items:
        if mail.Subject.startswith("JIRA"):  # Filtrer les mails par sujet si nécessaire
            try:
                # Exemple pour extraire le numéro de ticket et la compagnie
                # Adapter les regex en fonction du format des mails JIRA
                ticket_number = re.search(r"TICKET-\d+", mail.Body)
                company = re.search(r"Company: (\w+)", mail.Body)

                if ticket_number and company:
                    # Ajouter les informations extraites dans la liste
                    tickets_data.append({
                        "Ticket": ticket_number.group(),
                        "Company": company.group(1)
                    })

            except Exception as e:
                print(f"Erreur lors de l'analyse du mail : {e}")

    # Convertir les données en DataFrame pour l'exporter dans un fichier Excel
    if tickets_data:
        df = pd.DataFrame(tickets_data)
        df.to_excel("tickets_jira.xlsx", index=False)
        print("Données exportées avec succès dans tickets_jira.xlsx")
    else:
        print("Aucun mail de ticket JIRA trouvé ou aucun ticket/compagnie trouvé.")
else:
    print("Le dossier spécifié n'a pas été trouvé.")
