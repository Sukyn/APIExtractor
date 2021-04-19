import requests
import csv
import pandas as pd
import time
import json
import sys

# ----------------------------------------------------------------------------
# ---------------- INSTRUCTIONS POUR LE FICHIER REQUETES ---------------------
# ----------------------------------------------------------------------------
# on rentre les informations
# nécessaires à l'execution de l'API,
# - Les noms des comptes que l'on souhaite utiliser
# - Les jetons d'API associés
# - Les enquêtes pour lesquelles on souhaite exporter nos données
# ------------------------------------------------------------------------
# -------------------------------- MODELE --------------------------------
# ------------------------------------------------------------------------
#  {
#    "username": "exemple",
#    "token": "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
#    "lien": "https://admin-sphinx.u-psud.fr/",
#    "survey_names": [
#                    "EnqueteTresSpeciale2019", # Attention à ne pas
#                                               # oublier la virgule !
#                    "EnqueteMoinsSpeciale2021",
#                    "Enquete2017"              # Et pas de virgule après
#                                               # le dernier...
#                    ]
#  }, # On met aussi une virgule après notre bloc de compte

# Il faudra faire en sorte de réaliser une interface pour rentrer les
# noms des enquêtes qui permettrait d'éviter les erreurs de saisie


# On définit ici le nom des fichiers
# dans lesquels on va vouloir exporter
# nos tableaux
nom_fichier_emails = "data/statuts_"  # On ajoute le nom de l'enquête après
nom_fichier_resultats = "data/results_"  # On ajoute le nom de l'enquête après

# ----------------------------------------------------------------------------
# ---------------------------------- SCRIPT ----------------------------------
# ----------------------------------------------------------------------------
if __name__ == '__main__':

    # DEBUT DU SCRIPT
    print("Lancement du script, veuillez patienter...\n")

    # On enregistre l'heure de début d'execution
    # du script pour vérifier les performances
    start_time = time.time()

    # On récupère les requêtes de l'utilisateur
    with open("requetes.json") as f:
        data = json.load(f)

    # On va exécuter le script pour chaque compte renseigné
    for compte in data["requetes"]:

        # On annonce à l'utilisateur que l'on
        # commence à traiter ses demandes pour
        # ce compte là
        print(">>>>>>> Exports pour le compte " +
              compte["username"])

        # On garde ici les liens importants pour accéder à l'API
        # (ils peuvent changer si sphinx fait des mises à jour
        #  mais c'est normalement très rare !)
        version = "v4.0"
        api_link = compte["lien"] + "sphinxapi/api/" + version + "/"
        auth_link = compte["lien"] + "SphinxAuth/connect/token"

        # Le payload permet de définir les paramètres
        # que l'on enverra à l'API
        # Il permet notamment de vérifier notre identité grâce au jeton
        # et d'avoir accès aux données protégées par le système.
        payload = {
                   # Nom du compte
                   "username": compte["username"],
                   # Le token est notre clé de sécurité
                   # c'est ça qui permet de déverouiller les portes de l'API
                   # il est unique à chaque compte et doit rester privé
                   "token": compte["token"],
                   "lang": "FR",
                   "grant_type": "personal_token",
                   # Paramètre à ne pas modifier
                   "client_id": "sphinxapiclient"
                   }

        # On s'authentifie ici auprès de l'API, après une vérification
        # rapide de notre identité il nous offre un jeton
        # d'accès de 20 minutes à l'ensemble de ces données

        # Il faut regénérer ce jeton si on dépasse cette intervalle de temps
        # mais ça tombe bien, il n'y a pas besoin de modifier de paramètres
        # ça le fait tout seul
        try:
            clef = requests.post(auth_link,
                                 data=payload)
        except requests.exceptions.ConnectionError:
            print("La connexion ne s'établit pas bien, " +
                  "êtes-vous connecté au réseau ?")
            sys.exit()

        # On sauvegarde alors notre clé temporaire,
        # quand on fera des demandes à l'API on aura alors notre super
        # badge VIP sur nous, et on aura alors les accès !
        en_tete = {'Authorization': "bearer " + clef.json()["access_token"]}

        # On récupère la mailing list associée au compte
        # que l'on a définit au début du code
        # Et on n'oublie pas d'orner notre badge VIP ;)
        liste_emailing = requests.get(api_link + "mailing",
                                      headers=en_tete,
                                      data=payload)

        # On filtre les requêtes qui sont vides, mais qui sont plus pratiques
        # pour remplir le json
        compte["survey_names"] = [name for name
                                  in compte["survey_names"] if name != ""]

        # On va réaliser le script pour chaque compte
        # note : la variable id ici ne nous servira qu'à savoir
        # où l'on en est dans l'exécution du script
        for id, nom_enquete in enumerate(compte["survey_names"]):

            # On va rechercher alors la liste d'email
            # qui nous intéresse : celle de notre enquête
            mail_id = None
            for mail in liste_emailing.json():
                if (mail["associatedSurvey"] == nom_enquete):
                    mail_id = mail["mailingId"]
                    break

            # Si on en trouve pas, c'est pas normal !
            if mail_id is None:
                print("?? Le nom de l'enquête ne semble pas correct : " +
                      nom_enquete)
            else:
                # On récupère alors la mailing list (finalement !)
                emails = requests.get(api_link +
                                      f"mailing/{mail_id}/recipient",
                                      headers=en_tete,
                                      data=payload)


                # A ce stade, on a sauvegardé dans emails
                # les données de notre mailing list,
                # il ne nous restera plus qu'à les exploiter !

                # On fait pareil pour le résultat des enquêtes
                resultats = requests.get(api_link +
                                         f"survey/{nom_enquete}/data",
                                         headers=en_tete,
                                         data=payload)

                # On écrit dans un fichier la liste des informations
                # Pour cela on ouvre déjà le un fichier
                if emails.json():
                    with open(nom_fichier_emails +
                              f"{nom_enquete}.csv", "w") as csvfile:

                        # On ouvre le fichier en écriture
                        script = csv.writer(csvfile)

                        # On ne peut pas détailler davantage
                        # les informations personnelles car elles
                        # ne sont pas standardisées dans le fichier rendu
                        # par l'API
                        header = ['Adresse mail',
                                  'Statut',
                                  'Nb de contacts',
                                  'Date dernier contact']

                        for i in range(len(emails.json()[0]["params"].split(";"))):
                            header.append(f'info_{i}')
                        script.writerow(header)

                        # Et pour chaque email, on écrit les valeurs
                        for personne in emails.json():

                            # ON VERIFIE QUE LES CLEFS EXISTENT
                            # pour éviter les sauts de case
                            check_keys = ['recipientID',
                                          'address',
                                          'params',
                                          'status',
                                          'contactCount',
                                          'dateLastContact']
                            for key in check_keys:
                                if key not in personne:
                                    personne[key] = ''

                            # Module de traduction des statuts
                            # Note : l'appel à l'API en français
                            # renvoie quand même des variables
                            # anglaises pour les emails
                            translation = {
                                "Finished": "Fini",
                                "Contacted": "Contacté",
                                "Started": "Commencé",
                                "Blacklist": "Blacklist",
                                "Available": "Disponible",
                                "BadFormat": "Adresse Invalide"
                            }
                            personne["status"] = translation[personne["status"]]

                            # On récupère les valeurs de params
                            # qui sont spécifiques à l'enquête
                            info_specifiques = personne["params"].split(";")


                            personne.pop('params')

                            # Et on y ajoute les informations générales
                            # du diplômé
                            valeurs = list(personne.values())
                            valeurs.pop(0)
                            valeurs += info_specifiques

                            # On a alors une ligne avec les
                            # informations recherchées
                            script.writerow(valeurs)


                # On décode notre objet contenant les résultats
                # des enquêtes
                contenu = resultats.content.decode('utf-8')
                # et on coupe selon les points virgule
                valeurs = csv.reader(contenu.splitlines(), delimiter=';')

                # Et on écrit les données dans un fichier
                # data portant le nom de notre enquête

                # S'il y a des résultats, on les écrit dans le fichier
                resultats = list(valeurs)
                if resultats:
                    resultats.pop(0)
                    with open(nom_fichier_resultats +
                              f"{nom_enquete}.csv", "w") as csvfile2:

                        # On écrit les valeurs des réponses dans ce fichier
                        script = csv.writer(csvfile2)
                        script.writerows(resultats)

                # ON CONVERTIT LES DONNEES EN FICHIER XLSX
                # D'abord les résultats
                df_new = pd.read_csv(nom_fichier_resultats +
                          f"{nom_enquete}.csv")
                writer = pd.ExcelWriter(f"Resultats_{nom_enquete}.xlsx")
                df_new.to_excel(writer)
                writer.save()
                # Et les statuts des emails
                df_new = pd.read_csv(nom_fichier_emails +
                          f"{nom_enquete}.csv")
                writer = pd.ExcelWriter(f"Statuts_{nom_enquete}.xlsx")
                df_new.to_excel(writer, index=False)
                writer.save()



                # On envoie un message de log pour l'utilisateur...
                print("OK Enquête traitée : " +
                      nom_enquete +
                      " (Compte : " +
                      compte["username"] +
                      " " +
                      str(id+1) +
                      "/" +
                      str(len(compte["survey_names"])) +
                      ")")

    # FIN DU SCRIPT
    print("\nC'est fini ! Le script a pris " +
          str(round(time.time() - start_time, 2)) +
          " secondes")
