import smtplib
import json
import re

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os

class EnvoiPlanning:
    def __init__(self, window):
        self.window = window
        self.BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        self.PLANNINGS_PDF_DIR = os.path.join(self.BASE_DIR, "Data", "Plannings_pdf")
        self.EMPLOYEES_FILE = os.path.join(self.BASE_DIR, "Data", "Employes_json", "employees.json") # Chemin complet du fichier JSON
        self.MAILS_FILE = os.path.join(self.BASE_DIR, "mails.json")  # Chemin du fichier mails.json

        self.semaine_selected = self.get_semaine()

        self.dico_date_planning = {
            "semaine": "",
            "jours": "",
            "annee": "",
        }

    def get_semaine(self):
        return self.window.semaine_selected

    def get_donnees_employes(self):
        if os.path.exists(self.EMPLOYEES_FILE):
            with open(self.EMPLOYEES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return []

    def extraire_mails(self):
        return [
            {"email": employe.get("email", "")}
            for employe in self.get_donnees_employes()
            if employe.get("email")
        ]

    def get_dates_mails(self, semaine_selected):

        if semaine_selected and isinstance(semaine_selected, str): #semaine selected : "semaine XX - du DD/MM au DD/MM - YYYY"
            parties = re.split(" - ", semaine_selected)
            self.dico_date_planning["semaine"] = parties[0]
            print(self.dico_date_planning["semaine"])
            self.dico_date_planning["jours"] = parties[1]
            print(self.dico_date_planning["jours"])
            self.dico_date_planning["annee"] = parties[2]
            print(self.dico_date_planning["annee"])
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

    def send_email_with_pdf(self):
        # Obtenez les informations de configuration depuis mails.json
        self.get_dates_mails(self.semaine_selected)
        email_config = self.get_email_config()
        if email_config is None:
            return

        email_envoi = email_config.get("email", "")
        mdp = email_config.get("pwd", "")
        serveur = email_config.get("serveur", "")

        try:
            port = int(email_config.get("port", 465))
        except (ValueError, TypeError):
            print("Le port SMTP doit être un entier.")
            return
        #port = int(email_config.get("port", 465))

        if not email_envoi or not mdp or not serveur or not port:
            print("Les informations de configuration du serveur SMTP sont incomplètes.")
            return 0

        # Générer le nom du fichier PDF avec la semaine et l'année
        pdf_file_name = f"planning_{self.dico_date_planning['semaine'].replace(' ', '')}_{self.dico_date_planning['annee']}.pdf"  # Exemple: planning_semaine16_2025
        pdf_file_path = os.path.join(self.PLANNINGS_PDF_DIR, pdf_file_name)

        # Créer un message MIME multipart
        message = MIMEMultipart()
        message["From"] = email_envoi
        # Récupérer tous les e-mails des employés
        emails_destinataires = [employe["email"] for employe in self.extraire_mails()]
        email_vice_dir = "fleck.sophie@gmail.com"
        emails_destinataires.append(email_vice_dir)
        message["To"] = ", ".join(emails_destinataires)  # Placer tous les e-mails dans le champ "To"
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
                    smtp_server.sendmail(email_envoi, emails_destinataires, message.as_string())
                    print("E-mail envoyé avec succès avec la pièce jointe !")
            except Exception as e:
                print(f"Une erreur est survenue lors de l'envoi du mail : {e}")
        else:
            print(f"Le fichier PDF {pdf_file_name} n'existe pas à l'emplacement spécifié.")
