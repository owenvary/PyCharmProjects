import smtplib
import json
import re
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os

def resource_path(relative_path):
    """Retourne le chemin absolu d’un fichier compatible dev/exe"""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)
class EnvoiPlanning:
    def __init__(self, window):
        self.window = window

        self.PLANNINGS_PDF_DIR = resource_path("Data/Plannings_pdf")
        self.EMPLOYEES_FILE = resource_path("Data/Employes_json/employees.json")
        self.MAILS_FILE = resource_path("Data/Mails_json/mail_config.json")

        self.semaine_selected = self.get_semaine()

        self.dico_date_planning = {
            "semaine": "",
            "jours": "",
            "annee": "",
        }

    def get_semaine(self):
        # Récupère la semaine choisie
        return self.window.semaine_selected

    def get_donnees_employes(self):
        # Charge le fichier employees.json
        if os.path.exists(self.EMPLOYEES_FILE):
            with open(self.EMPLOYEES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return []

    def extraire_mails(self):
        # Récupérer les emails destinataires
        return [
            {"email": employe.get("email", "")}
            for employe in self.get_donnees_employes()
            if employe.get("email")
        ]

    def get_dates_mails(self, semaine_selected):
        # Récupérer les valeurs temporelles (DD/MM/YYYY) et les placent dans un dico
        if semaine_selected and isinstance(semaine_selected, str): #semaine selected : "semaine XX - du DD/MM au DD/MM - YYYY"
            parties = re.split(" - ", semaine_selected)
            self.dico_date_planning["semaine"] = parties[0]
            #print(self.dico_date_planning["semaine"])
            self.dico_date_planning["jours"] = parties[1]
            #print(self.dico_date_planning["jours"])
            self.dico_date_planning["annee"] = parties[2]
            #print(self.dico_date_planning["annee"])
            return
        else:
            return "Erreur lors de la lecture de la semaine choisie"

    def get_email_config(self):
        # Charger la configuration du fichier mails.json
        if os.path.exists(self.MAILS_FILE):
            with open(self.MAILS_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
                return config
        else:
            print("Le fichier mails.json est manquant ou invalide.")
            return None

    def send_email_with_pdf(self, selected_names=None):
        """
        Envoi le mail aux personnes sélectionnés, l'objet, le message et la pièce jointe sont dynamiques.
        """
        employes = self.get_donnees_employes()

        # Créer un dictionnaire nom -> email
        nom_to_email = {employe.get("nom", ""): employe.get("email", "") for employe in employes}

        if selected_names is None:
            # Aucun nom sélectionné -> on prend tout le monde
            selected_emails = [email for email in nom_to_email.values() if email]
        else:
            # Retrouver les emails correspondant aux noms sélectionnés
            selected_emails = [nom_to_email.get(nom) for nom in selected_names if nom_to_email.get(nom)]

        # Ajouter Codir si besoin
        codir_email = "fleck.sophie@gmail.com"
        if "Sophie" in selected_names and codir_email not in selected_emails:
            selected_emails.append(codir_email)

        # Vérifier l'existence de mails.json
        self.get_dates_mails(self.semaine_selected)
        email_config = self.get_email_config()
        if email_config is None:
            return

        # Récupère les champs de mails.json
        email_envoi = email_config.get("email_envoi", "")
        mdp = email_config.get("pwd", "")
        serveur = email_config.get("serveur", "")

        try:
            port = int(email_config.get("port", 465))
        except (ValueError, TypeError):
            print("Le port SMTP doit être un entier.")
            return

        if not email_envoi or not mdp or not serveur or not port:
            print("Les informations de configuration du serveur SMTP sont incomplètes.")
            return 0

        # Générer le nom du fichier PDF avec la semaine et l'année
        pdf_file_name = f"planning_{self.dico_date_planning['semaine'].replace(' ', '')}_{self.dico_date_planning['annee']}.pdf"  # Exemple: planning_semaine16_2025
        pdf_file_path = os.path.join(self.PLANNINGS_PDF_DIR, pdf_file_name)

        # Créer un message MIME multipart
        message = MIMEMultipart()
        message["From"] = email_envoi
        message["To"] = ", ".join(selected_emails)  # Utiliser les emails sélectionnés
        message["Subject"] = f"Planning de la {self.dico_date_planning['semaine']}"


        # Texte du message avec les informations de la semaine
        mail_pro = "patrice_vary@franchise.carrefour.com"
        tel_pro = "0625538184"
        message_html = f"""
        <html>
          <body>
            <p>Bonjour,</p><br>
            <p>
              Veuillez trouver ci-joint le planning de la <strong>{self.dico_date_planning['semaine']}</strong>
              pour la période <strong>{self.dico_date_planning['jours']}</strong> de l'année
              <strong>{self.dico_date_planning['annee']}</strong>.
            </p><br>
            <p>Cordialement,<br>
            Patrice VARY</p>
            <hr>
            <p><em>Merci de ne pas répondre à ce mail. Pour toute question, veuillez envoyer un mail à l'adresse suivante :</em>
            <strong> {mail_pro}</strong>,<br>
            <em>ou contacter le</em> <strong>{tel_pro}</strong>.</p>
          </body>
        </html>
        """

        message.attach(MIMEText(message_html, "html"))


        # Vérifier si le fichier PDF existe avant de l'envoyer
        if os.path.exists(pdf_file_path):
            try:
                # Joindre le fichier PDF à l'email
                with open(pdf_file_path, "rb") as pdf_file:
                    pdf_part = MIMEApplication(pdf_file.read(), _subtype="pdf")
                    pdf_part.add_header('Content-Disposition', 'attachment', filename=pdf_file_name)
                    message.attach(pdf_part)

                # Connexion au serveur SMTP de Gmail et envoi de l'e-mail
                with smtplib.SMTP_SSL(serveur, port) as smtp_server:
                    smtp_server.login(email_envoi, mdp)
                    smtp_server.sendmail(email_envoi, selected_emails, message.as_string())
                    print("E-mail envoyé avec succès avec la pièce jointe !")
                    return True #Pour que success = True et affiché le bon popup
            except Exception as e:
                print(f"Une erreur est survenue lors de l'envoi du mail : {e}")
        else:
            print(f"Le fichier PDF {pdf_file_name} n'existe pas à l'emplacement spécifié.")

