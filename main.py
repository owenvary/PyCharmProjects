import sys
import json
import os
import locale
import re
import webbrowser
from collections import defaultdict


from PySide6.QtGui import QFont, QBrush, QColor, QKeyEvent, QKeySequence, QShortcut, QIcon, QGuiApplication
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTableWidget, QLabel, QPushButton,
    QGridLayout, QTableWidgetItem, QHeaderView, QSizePolicy,
    QComboBox, QVBoxLayout, QHBoxLayout, QMessageBox, QDialog, QAbstractItemView,
)


from PySide6.QtCore import Qt, QSize
from datetime import datetime, timedelta, date
from PIL import Image

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

from gestion_employes import GestionEmployes
from envoi_mails import EnvoiPlanning
from selection_mails import SelectionMails


def resource_path(relative_path):
    """Retourne le chemin absolu d’un fichier compatible dev/exe"""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)

class Window(QMainWindow):
    def __init__(self):
        super().__init__()

        # Import des données/ Création des rep.
        self.PLANNINGS_JSON_DIR = resource_path("Data/Plannings_JSON")
        self.PLANNINGS_PDF_DIR = resource_path("Data/Plannings_PDF")
        os.makedirs(self.PLANNINGS_JSON_DIR, exist_ok=True)
        os.makedirs(self.PLANNINGS_PDF_DIR, exist_ok=True)

        # Guide d'utilisation -> pré-version, à upgrade avant déploiement
        self.guide_utilisation = resource_path("pré-Guide d'utilisation FastPlanning.pdf")
        # Icônes des boutons
        self.save_icon = resource_path("Icones/save_icon.png")
        self.send_icon = resource_path("Icones/send_icon.png")
        self.load_icon = resource_path("Icones/load_icon.png")
        self.edit_icon = resource_path("Icones/edit_icon.png")
        self.analyze_icon = resource_path("Icones/analyze_icon.png")
        self.info_icon = resource_path("Icones/info_icon.png")
        # Logo de l'app
        self.fast_planning_img = resource_path("Images/FastPlanning_logo.png")
        # Configuration générale
        locale.setlocale(locale.LC_TIME, 'fr_FR')  # Pour afficher les dates en français
        # Paramétrage de la fenêtre principale
        self.setWindowTitle("FastPlanning")
        self.setWindowState(Qt.WindowMaximized) # Pleiné cran
        self.setWindowIcon(QIcon(self.fast_planning_img)) # Logo -> icône de la fenêtre (barre de titre)

        # Initialisation de l'objet pour la gestion des employés
        self.gestion_employes = GestionEmployes()

        # Chargement des employés depuis un fichier
        self.load_employees() # Charge le json modifié depuis gestion_employes
        self.nb_employees = len(self.employees) # =nb de lignes du planning

        # Panneau central de l'interface
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout(self.central_widget)

        # Panneau du haut (sélection de la semaine et boutons "Charger" et "Employés")
        top_layout = QHBoxLayout()

        # Menu déroulant pour la sélection de la semaine
        self.selection_semaines = QComboBox()
        self.selection_semaines.setMinimumWidth(350)
        self.initialiser_options_semaines()

        # Bouton pour charger le planning
        self.btn_charger = QPushButton()
        self.btn_charger.setIcon(QIcon(self.load_icon)) # png -> icône
        self.btn_charger.setIconSize(QSize(48, 48)) # Taille de l'icône
        self.btn_charger.setToolTip("Charger planning")

        # Bouton pour ouvrir le menu de gestion
        self.btn_employes = QPushButton("Gestion du personnel")
        self.btn_employes.setIcon(QIcon(self.edit_icon))
        self.btn_employes.setIconSize(QSize(48, 48))
        self.btn_employes.setToolTip("Modifier le personnel")
        self.btn_employes.clicked.connect(self.lancement_interface_employes)

        # ---TOP LAYOUT---
        # Assemblage du menu déroulanbt et du btn charger
        left_layout = QHBoxLayout()
        left_layout.addWidget(self.selection_semaines)
        left_layout.addSpacing(15)
        left_layout.addWidget(self.btn_charger)

        # Assemblage du layout de gauche avec le menu de gestion
        top_layout.addLayout(left_layout)
        top_layout.addStretch() # ajoute un espace
        top_layout.addWidget(self.btn_employes)

        # ---CENTER LAYOUT--- Tableau 2x1 (titre \n table planning)
        center_layout = QGridLayout()

        #Titre dynamique selon la semaine choisie - initialisé à la sem. actuelle
        self.semaine_selected = self.selection_semaines.currentText()
        self.selection_semaines.currentIndexChanged.connect(self.semaine_choisie)
        self.label_title = QLabel(f"Planning - {self.semaine_selected}")
        self.label_title.setFont(QFont('Segoe UI', 16))


        # Grille de planning
        self.grille_planning = QTableWidget(self.nb_employees, 9)
        self.maj_headers_planning()

        # Configuration flexible selon le contenu
        self.grille_planning.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.grille_planning.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)# Initialisation de la colonne des employés

        # MAJ des la première colonne selon employees.json
        self.update_employes_column()



        # Sélection multiple dans la grille (permet les Ctrl C - V - X - Z - Y)
        self.grille_planning.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self.grille_planning.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.grille_planning.installEventFilter(self)

        # Assemblage du titre et du planning
        center_layout.addWidget(self.label_title, 0, 0, alignment=Qt.AlignCenter)
        center_layout.addWidget(self.grille_planning, 1, 0)

        # Application des styles à la grille
        self.apply_headers_style() # Style des headers du planning
        self.apply_table_style() # Style des cellules du planning
        self.apply_font_to_table() # Police du texte dans les cellules
        self.apply_row_colors() # Couleur de backg. du planning (alternance)

        # ---BOTTOM LAYOUT--- Boutons Infos, Analyser, Enregistrer et Envoyer/mail
        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(40)

        # -- BOUTONS --
        # Initialisation
        self.btn_info = QPushButton()
        self.btn_analyser = QPushButton()
        self.btn_sauvegarder = QPushButton()
        self.btn_envoyer = QPushButton()

        # Icônes
        self.btn_info.setIcon(QIcon(self.info_icon))
        self.btn_analyser.setIcon(QIcon(self.analyze_icon))
        self.btn_sauvegarder.setIcon(QIcon(self.save_icon))
        self.btn_envoyer.setIcon(QIcon(self.send_icon))

        # ToolTips
        self.btn_info.setToolTip("Ouvrir le guide d'utilisation")
        self.btn_analyser.setToolTip("Analyser le planning")
        self.btn_sauvegarder.setToolTip("Enregistrer/Sauvegarder planning")
        self.btn_envoyer.setToolTip("Envoyer mails")

        # Style pour les boutons
        style_btn = """
        QPushButton {
            color: black;
            background-color: #e0e0e0;  /* léger gris clair au survol */
            border-radius: 6px;
            padding: 15px;
        }
        QPushButton:hover {
            background-color: #007A33;  /* vert foncé */
            border-radius: 6px;
        }
        QPushButton:pressed {
        background-color: #005F26;  /* Vert encore plus foncé lors du clic */
        }
        """

        # Application du style à tous les boutons
        for btn in (
        self.selection_semaines, self.btn_charger, self.btn_employes, self.btn_sauvegarder, self.btn_envoyer, self.btn_analyser, self.btn_info):
            btn.setFont(QFont('Segoe UI', 14)) # Police & taille
            btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed) # Verouillage malgré la taille de la fenêtre
            btn.setIconSize(QSize(48, 48)) # Taille icône
            btn.setStyleSheet(style_btn) # Appliquer la style sheet

        # Assemblage du bas de l'interface - 4 btn centrés
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.btn_info)
        bottom_layout.addWidget(self.btn_analyser)
        bottom_layout.addWidget(self.btn_sauvegarder)
        bottom_layout.addWidget(self.btn_envoyer)
        bottom_layout.addStretch()

        # Cells completion
        self.grille_planning.cellChanged.connect(self.eventListener_chgt_cellule)

        # Connexion des boutons aux fonctions
        self.btn_charger.clicked.connect(self.charger_planning)
        self.btn_sauvegarder.clicked.connect(self.enregistrer_planning)
        self.btn_envoyer.clicked.connect(self.envoi_planning)
        self.btn_analyser.clicked.connect(self.analyser_planning)
        self.btn_info.clicked.connect(self.ouvrir_guide)

        # Assemblage des 3 layouts
        main_layout.addLayout(top_layout)
        main_layout.addLayout(center_layout)
        main_layout.addLayout(bottom_layout)

        # Initialisation des stacks pour l'historique et la gestion des actions (Ctrl+Z / Ctrl+Y)
        self.historique = defaultdict(list)
        self.redo_stack = defaultdict(list)

        # Initialise une stack pour chaque tuple (row, col)
        self.init_historique()

        #clipboard
        self.clipboard_text = QGuiApplication.clipboard().text()

        # Raccourcis clavier
        # Re-do / do
        self.raccourci_ctrl_z = QShortcut(QKeySequence("Ctrl+Z"), self)
        self.raccourci_ctrl_z.activated.connect(self.retour_en_arriere)
        self.raccourci_ctrl_y = QShortcut(QKeySequence("Ctrl+Y"), self)
        self.raccourci_ctrl_y.activated.connect(self.refaire)
        # Save
        self.raccourci_ctrl_s = QShortcut(QKeySequence("Ctrl+S"), self)
        self.raccourci_ctrl_s.activated.connect(self.enregistrer_planning)
        # Couper
        self.raccourci_ctrl_x = QShortcut(QKeySequence("Ctrl+X"), self)
        self.raccourci_ctrl_x.activated.connect(self.couper_cellules_selectionnees)

    # ---BTN Listeners---

    def charger_planning(self):
        """Charge un planning depuis un fichier JSON, en fonction de la semaine sélectionnée."""

        # Récupérer la semaine sélectionnée et l'année courante
        semaine_num = self.selection_semaines.currentData()
        annee = datetime.now().year

        # Chemin du fichier JSON à charger
        filename = resource_path(f"Data/Plannings_json/planning_semaine{semaine_num}_{annee}.json")

        # Vérifier que le fichier existe
        if not os.path.exists(filename):
            QMessageBox.warning(self, "Fichier introuvable", f"Aucun planning trouvé pour la semaine {semaine_num}.")
            return

        # Charger les données depuis le fichier
        with open(filename, "r", encoding="utf-8") as f:
            planning_data = json.load(f)

        # Nettoyer toutes les anciennes données (hors colonne Nom et Total)
        for row in range(self.grille_planning.rowCount()):
            for col in range(1, self.grille_planning.columnCount() - 1):
                self.grille_planning.setItem(row, col, QTableWidgetItem(""))

        # Liste des jours de la semaine (alignée avec les colonnes 1 à 7)
        jours = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]

        # Création d’un dictionnaire pour retrouver rapidement la ligne d’un employé à partir de son nom
        nom_to_row = {}
        for row in range(self.grille_planning.rowCount()):
            item = self.grille_planning.item(row, 0)
            if item:
                nom_to_row[item.text().strip().lower()] = row

        # Si le nom est reconnu, charger la ligne, sinon l'initialiser comme vide
        for emp_data in planning_data.get("employes", []):
            nom = emp_data.get("nom", "").strip().lower()
            if nom in nom_to_row:
                row = nom_to_row[nom]
                horaires = emp_data.get("horaires", {})
                for col, jour in enumerate(jours):
                    texte = horaires.get(jour, "")
                    self.grille_planning.setItem(row, col + 1, QTableWidgetItem(texte))

        # Rafraîchir les styles après chargement (éviter bug graphiques)
        self.apply_font_to_table()
        self.apply_row_colors()

        # Pop-ups
        QMessageBox.information(self, "Succès", f"planning_semaine{semaine_num}_{annee} a été chargé avec succès")
        print(f"Planning de la semaine {semaine_num} chargé avec succès.")
    def ouvrir_guide(self):
        """
        Ouvre un guide d'utilisationd e l'application pdf. Cette fonction se contente d'ouvrir un path dans le nav.
        """
        # Vérifier si le fichier existe
        if os.path.exists(self.guide_utilisation):
            # Ouvrir le fichier PDF dans le navigateur par défaut
            webbrowser.open(f'file:///{os.path.abspath(self.guide_utilisation)}')
        else:
            # Si le fichier n'existe pas, afficher un message d'erreur
            print("Le fichier PDF du guide d'utilisation n'existe pas.")
            QMessageBox.warning(self, "Erreur", "Le fichier du guide d'utilisation est introuvable.")
    def analyser_planning(self):
        """
        Fait apparaître un pop-up analysant les sous-effectifs, les chevauchements et les totaux du plannings.
        """
        erreurs = []
        jours = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
        total_col = self.grille_planning.columnCount() - 1  # La dernière colonne = total

        # Adapter les horaires pour chaque jour
        presence_par_jour = {
            jour: {h: 0 for h in
                   self.generate_hours(8.0 if jour == "dimanche" else 7.0, 13.25 if jour == "dimanche" else 21.25)}
            for jour in jours
        }

        for row in range(self.grille_planning.rowCount()):
            nom_employe = self.grille_planning.item(row, 0).text().strip()

            # Trouver le contrat de l'employé dans self.employees
            contrat = ""
            for emp in self.employees:
                if emp["nom"] == nom_employe:
                    contrat = emp.get("contrat", "")
                    break

            # Récupérer les heures déjà calculées dans la dernière colonne
            total_item = self.grille_planning.item(row, total_col)
            try:
                total_heures = float(total_item.text().replace(",", "."))
            except (ValueError, AttributeError):
                total_heures = 0

            # Analyser chevauchements et présence
            for col in range(1, total_col):
                item = self.grille_planning.item(row, col)
                if item is None or not item.text().strip():
                    continue

                jour = jours[col - 1]
                horaires = self.extraire_creneaux(item.text().strip())

                # Chevauchements
                if any(self.chevauche(h1, h2) for i, h1 in enumerate(horaires) for h2 in horaires[i + 1:]):
                    erreurs.append(f"• Chevauchement d'horaires, {nom_employe} le {jour}")

                # Présences
                for debut, fin in horaires:
                    # Si l'heure de début et de fin sont identiques, ignorer
                    if debut == fin:
                        continue

                    h = debut
                    while h < fin:
                        if h in presence_par_jour[jour]:
                            presence_par_jour[jour][h] += 1
                        h += 1 if h + 1 <= 21 else 0.25

            # Comparaison au contrat
            if contrat not in ("ÉTUDIANT", "PATRON"):
                contrat_heures = {
                    "33h": 33,
                    "35h": 35,
                    "39h": 39,
                    "ALTERNANT": 35,
                }.get(contrat, 0)

                if contrat_heures > 0:
                    difference = total_heures - contrat_heures
                    if difference != 0:
                        surplus_ou_manque = "Heures sup" if difference > 0 else "Manque d'h"
                        erreurs.append(
                            f"• {nom_employe} - {surplus_ou_manque}: "
                            f"{abs(difference):.2f}h (contrat: {contrat_heures}h)"
                        )

        # Sous-effectif
        for jour, heures in presence_par_jour.items():
            debut = None
            for h in sorted(heures):
                # Ne pas signaler un sous-effectif si l'heure de début = fin
                if heures[h] < 3 and debut is None:  # Le sous-effectif commence
                    debut = h
                elif debut is not None and (heures[h] >= 3 or h == max(heures)):  # Fin du sous-effectif
                    if debut != h:  # Ne pas signaler si c'est la même heure de début et de fin
                        erreurs.append(
                            f"• Sous-effectif le {jour} de {self.h_float_to_str(debut)} à {self.h_float_to_str(h)}")
                    debut = None
            if debut is not None and debut != max(heures):
                erreurs.append(
                    f"• Sous-effectif le {jour} de {self.h_float_to_str(debut)} à {self.h_float_to_str(max(heures))}"
                )

        # Résultat final
        QMessageBox.information(self, "Analyse du planning", "\n".join(erreurs) if erreurs else "Aucune erreur détectée. ✅")


    def enregistrer_planning(self):
        """
        Enregistre le planning en pdf pour l'envoi et en json pour la charge.
        """
        # Demander confirmation avant d'envoyer l'email
        confirm = QMessageBox.question(
            self, "Confirmation", "Voulez-vous enregistrer le planning ?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes,
        )
        if confirm == QMessageBox.Yes:
            self.enregistrer_pdf()  # Sauvegarde PDF pour l'envoi et l'impression
            self.enregistrer_json()  # Sauvegarde JSON pour le chargement
        elif confirm == QMessageBox.No:
            QMessageBox.information(self, "Annulé", "L'enregistrement du planning a été annulé")

    def envoi_planning(self):
        # Demander confirmation avant d'envoyer l'email
        confirm = QMessageBox.question(
            self, "Confirmation", "Voulez-vous envoyer le planning ?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if confirm == QMessageBox.Yes:
            # Obtenir la liste des employés
            employes = EnvoiPlanning(self).get_donnees_employes()

            # Créer une fenêtre de sélection d'employés
            dialog = SelectionMails(employes, self)
            if dialog.exec() == QDialog.Accepted:
                selected_emails = dialog.get_selected_employees()

                if selected_emails:
                    # Si l'utilisateur a sélectionné des employés, envoyer l'email
                    envoi_planning = EnvoiPlanning(self)
                    envoi_planning.send_email_with_pdf(selected_emails)

                    # Affichage du pop-up de confirmation
                    QMessageBox.information(self, "Succès",
                                            "Le planning a été envoyé avec succès aux employés sélectionnés.")
                else:
                    # Aucun employé sélectionné
                    QMessageBox.warning(self, "Aucun sélectionné", "Aucun employé n'a été sélectionné pour l'envoi.")
            else:
                QMessageBox.information(self, "Annulé", "L'envoi du planning a été annulé.")
        else:
            QMessageBox.information(self, "Annulé", "L'envoi du planning a été annulé.")

    # ---Outils---
        # Manip de données
    def generate_hours(self, start, end):
        """
        Retournes une liste découpant les créneaux horaires par pas de 1h, 15min si c'est la dernière heure.
        """
        heures = []
        h = start
        while h + 1 <= end:
            heures.append(round(h, 2))
            h += 1
        if round(end - int(end), 2) == 0.25:
            heures.append(round(end, 2))
        return heures

    def chevauche(self, c1, c2):
        return max(c1[0], c2[0]) < min(c1[1], c2[1])

    def extraire_creneaux(self, texte):
        """
        Extrait les creneaux "7 - 14" en [7, 14] pour pouvoir les manipuler.
        """
        try:
            heures = list(map(float, re.split(r"[-\s]+", texte)))
            return [(heures[i], heures[i + 1]) for i in range(0, len(heures) - 1, 2)]
        except Exception:
            return []

    def h_float_to_str(self, h):
        """Convertis les heures float (7.25h en 7h15)"""
        h_int = int(h)
        m = int(round((h - h_int) * 60))
        return f"{h_int:02d}h{m:02d}"

        #Evnmts
    def init_historique(self):
        """Initialise l'historique pour toutes les cellules de la table."""
        for row in range(self.grille_planning.rowCount()):
            for col in range(self.grille_planning.columnCount()):
                item = self.grille_planning.item(row, col)
                if not item:
                    item = QTableWidgetItem("")  # Initialiser la cellule si elle est vide
                    self.grille_planning.setItem(row, col, item)
                # Initialiser l'historique avec le texte de la cellule au début
                self.historique[(row, col)] = [item.text()]
                self.redo_stack[(row, col)] = []  # Initialisation du stack de redo

    def mettre_a_jour_historique_et_total(self, row, col):
        """
        Met à jour l'historique de la cellule (row, col) puis met à jour le total de la ligne correspondante.
        """
        item = self.grille_planning.item(row, col)
        if item:
            texte = item.text().strip()
            key = (row, col)

            # Mise à jour de l'historique de la cellule
            if key not in self.historique:
                self.historique[key] = []
            if not self.historique[key] or self.historique[key][-1] != texte:
                self.historique[key].append(texte)
                self.redo_stack[key] = []

            # Calcul du total de la ligne
            self.calculer_total_ligne(row)

    def keyPressEvent(self, event):
        """
        Déclenche des fonctions selon une touche pressé, ici suppression
        """
        # Effacer des cellules sélectionnées
        if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            self.effacer_cellules_selectionnees()
        current_item = self.grille_planning.currentItem()
        if current_item:
            row = self.grille_planning.currentRow()
            col = self.grille_planning.currentColumn()
            texte_courant = current_item.text()
            key = (row, col)
            # MAJ stacks
            if not self.historique[key] or texte_courant != self.historique[key][-1]:
                self.historique[key].append(texte_courant)
                self.redo_stack[key] = []

        super().keyPressEvent(event)

    def eventFilter(self, source, event):
        """
        Déclenche les fonctions copier-collers en fonction des signaux détectés
        """
        if source == self.grille_planning and isinstance(event, QKeyEvent):
            if event.matches(QKeySequence.Copy):
                self.copier_cellules_selectionnees()
                return True
            elif event.matches(QKeySequence.Paste):
                self.grille_planning.blockSignals(True)
                self.coller_cellules_selectionnees()
                self.grille_planning.blockSignals(False)
                return True
        return super().eventFilter(source, event)

    # Remplissage des cellules
    def eventListener_chgt_cellule(self, row, col):
        """Transforme les espaces en ' - ' pour les colonnes de planning uniquement (1 à 8).
        Enregistre uniquement le texte final transformé dans l'historique.
        """
        # Ne rien faire si la colonne est 0 (noms) ou hors des colonnes de planning
        if col == 0 or col > 8:
            return

        item = self.grille_planning.item(row, col)
        if item:
            original_text = item.text()

            # Ne rien faire si le texte contient déjà un tiret
            if "-" in original_text:
                return

            result_text = ""
            space_counter = 0

            for char in original_text:
                if char == " ":
                    space_counter += 1
                    if (space_counter % 2) != 0:
                        result_text += " - "
                    else:
                        result_text += "   "
                else:
                    result_text += char

            # Appliquer la modification uniquement si nécessaire
            if result_text != original_text:
                self.grille_planning.blockSignals(True)
                item.setText(result_text)
                self.grille_planning.blockSignals(False)

            # Enregistrer la version transformée (finale) dans l'historique
            final_text = item.text()
            if isinstance(self.historique[(row, col)], list):
                if not self.historique[(row, col)] or final_text != self.historique[(row, col)][-1]:
                    self.historique[(row, col)].append(final_text)
                    self.redo_stack[(row, col)] = []
            else:
                self.historique[(row, col)] = [final_text]
                self.redo_stack[(row, col)] = []

            self.calculer_total_ligne(row)

        # Grille Planning
    def set_planning_title(self):
        self.label_title.setText("Planning - " + self.semaine_selected)

    def maj_headers_planning(self):
        """ Récupère la date du premier jour de la semaine choisie puis calcule les jours suivants afin de maj les headers"""
        start_of_week = self.get_start_of_week()
        week_days = [start_of_week + timedelta(days=i) for i in range(7)]

        headers = [""] # Pas de nom pour la colonne employé
        for day in week_days:
            day_name = day.strftime('%A').capitalize()
            headers.append(f"{day_name} {day.strftime('%d/%m')}")
        headers.append("Total") # Total est le dernier header
        self.grille_planning.setHorizontalHeaderLabels(headers)

    def calculer_total_ligne(self, row):
        """
        Fait la somme des créneaux d'une ligne et l'imprime dans la case 'Total' correspondante.
        """
        total = 0.0
        total_col = self.grille_planning.columnCount() - 1
        nom_item = self.grille_planning.item(row, 0)
        nom_employe = nom_item.text().strip() if nom_item else ""

        # Récupérer le contrat associé à l'employé
        contrat = ""
        for emp in self.employees:
            if emp["nom"] == nom_employe:
                contrat = emp.get("contrat", "")
                break
        # Convertir le contrat en entier (h/semaines)
        if contrat != "PATRON":
            contrat_value = {
                "33h": 33,
                "35h": 35,
                "39h": 39,
                "ALTERNANT": 35,
                "ÉTUDIANT": 35,
            }.get(contrat, 0)
        else:
            contrat_value = 0
        # Convertir les champs (*str) en entiers
        for col in range(total_col):
            item = self.grille_planning.item(row, col)
            if item:
                # Récupérer le texte dans la cellule et le lire comme étant en majuscules
                texte = item.text().strip()
                texte = texte.upper()
                if texte:
                    try:
                        if texte in ("CP", "CGP", "CONGÉS"):
                            # Congé payé = 1 jour = contrat hebdo / 5
                            if contrat_value:
                                total += contrat_value / 6
                        elif texte in ("AFORMANCE", "CFA", "COURS", "ALTERNANCE"):
                            total += 7
                        elif isinstance(texte, str):
                            parties = re.split(r"[-\s]+", texte)
                            heures = list(map(float, parties))
                            for i in range(0, len(heures) - 1, 2):
                                debut = heures[i]
                                fin = heures[i + 1]
                                total += abs(fin - debut)
                        # Imprimer l'erreur dans la console mais l'ignorer dans l'app.
                        else:
                            print(f"Texte invalide dans la cellule ({row}, {col}): {texte}")
                    except ValueError:
                        pass

        total_item = self.grille_planning.item(row, total_col)
        if total_item is None:
            total_item = QTableWidgetItem(f"{round(total, 2)}")  # Arrondis à 2 décimales
            total_item.setFlags(Qt.ItemIsEnabled)
            self.grille_planning.setItem(row, total_col, total_item)  # Imprimer Total de la ligne
        else:
            total_item.setText(f"{round(total, 2)}")

    def update_employes_column(self):
        """
        Initialise la colonne "Nom" et permet aussi de la MAJ après modif.
        """
        # Récup les noms
        noms = [emp["nom"] for emp in self.employees]
        # Set up le nb de lignes
        self.grille_planning.setRowCount(len(noms))

        for row in range(len(noms)):
            item = self.grille_planning.item(row, 0)
            if item is None:
                item = QTableWidgetItem()
                self.grille_planning.setItem(row, 0, item)

            item.setText(noms[row])
            item.setFont(QFont('Segoe UI', 14))
            item.setTextAlignment(Qt.AlignCenter)  # Au milieu de la case

            # Rendre la cellule non éditable afin de conserver ses propriétés
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)

        # Re-appliquer le design des cellules ayant été modifié par les selections
        self.apply_row_colors()
        self.apply_font_to_table()


        # Datas
    def load_employees(self):
        """
        Charge le fichier employees.json
        """
        employees_file = resource_path("Data/Employes_json/employees.json")
        if os.path.exists(employees_file):
            with open(employees_file, "r", encoding="utf-8") as f:
                self.employees = json.load(f)
        else:
            self.employees = []
            self.update_employes_column()

    def semaine_choisie(self, index):
        """ met à jour le planning et le titre selon la semaine choisie
        :param index: numéro de la semaine choisie"""
        self.semaine_selected = self.selection_semaines.itemText(index)
        self.set_planning_title()
        self.maj_headers_planning()

    def get_start_of_week(self):
        """
        Retourne le premier jour (lundi) de la semaine sélectionnée dans le menu déroulant.
        Si aucune semaine n'est sélectionnée, retourne la date du jour.
        """
        if self.selection_semaines.currentData() is None:
            return date.today()

        week_number = int(self.selection_semaines.currentData())  # Numéro de semaine (1 à 52)
        return date.fromisocalendar(date.today().year, week_number, 1)  # 1 = Lundi

    def initialiser_options_semaines(self):
        """Intégrer les semaines du menu déroulant de sélection si comprise dans l'intervalle [-3mois, J actuel, +1 mois]"""
        # Récupérer J - M - A
        year = date.today().year
        today = date.today()
        week = 1
        # Récupérer le numéro de semaine actuel
        current_week = today.isocalendar()[1]  # Récupère la semaine actuelle

        while True:
            try:
                monday = date.fromisocalendar(year, week, 1)
                sunday = monday + timedelta(days=6)
                # Si le début de la semaine i est à moins de 3 mois avant le jour actuel on intègre la semaine
                if monday < today and (today - monday).days <= 90:
                    item_text = f"Semaine {week} - du {monday.strftime('%d/%m')} au {sunday.strftime('%d/%m')} - {year}"
                    self.selection_semaines.addItem(item_text, userData=week)
                # Si le début de semaine à moins de 1 mois avant le jour actuel on l'intègre
                elif monday >= today and (monday - today).days <= 30:
                    item_text = f"Semaine {week} - du {monday.strftime('%d/%m')} au {sunday.strftime('%d/%m')} - {year}"
                    self.selection_semaines.addItem(item_text, userData=week)
                # Si + 30j on sort de la boucle
                if monday > today + timedelta(days=30):
                    break
                week += 1
            except ValueError:
                break

        # Trouver et sélectionner la semaine actuelle
        for i in range(self.selection_semaines.count()):
            item_text = self.selection_semaines.itemText(i)
            if f"Semaine {current_week}" in item_text:
                self.selection_semaines.setCurrentIndex(i)
                break

    def enregistrer_pdf(self):
        """
        Enregistre et génère le planning au format pdf
        """
        # Données temporelles pour sauvegarder avec le bon nom de dficheir
        semaine_num = self.selection_semaines.currentData()
        annee = datetime.now().year
        date_debut, date_fin = self.get_dates_semaine(semaine_num)  # Tu dois avoir une fonction ou méthode pour ça
        pdf_filename = os.path.join(self.PLANNINGS_PDF_DIR, f"planning_semaine{semaine_num}_{annee}.pdf")

        # Paramètre de la page (marges et format paysage)
        doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(A4), leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
        elements = []

        # HEADER: Logo Carrefour + Titre
        logo_path = resource_path("Images/logo_carrefour_city.png")

        try:
            logo = Image(logo_path, width=120, height=160)
        except:
            logo = Paragraph("Logo manquant", getSampleStyleSheet()["Normal"])

        # Enregistrer une nouvelle police
        try:
            pdfmetrics.registerFont(TTFont("Montserrat-Bold", resource_path("Police/Montserrat-Bold.ttf")))
            font_title = "Montserrat-Bold"

        except:
            font_title = "Helvetica-Bold"

        # Titre à droite du logo
        title_style = ParagraphStyle(
            name="TitreCarrefour",
            fontName=font_title,
            fontSize=20,
            alignment=TA_RIGHT,
            textColor=colors.HexColor("#007A33")
        )
        title = Paragraph("Carrefour City – Saint-Brieuc", title_style)

        # Logo + titre côte à côte
        header_table = Table([[logo, title]], colWidths=[100, 600])
        header_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        elements.append(header_table)
        elements.append(Spacer(1, 10))

        # Titre du planning : Semaine X du ... au ... ===
        planning_title_style = ParagraphStyle(
            name="TitrePlanning",
            fontName="Helvetica-Bold",
            fontSize=14,
            alignment=TA_CENTER,
            textColor=colors.black
        )
        titre_planning = f"Planning – Semaine {semaine_num} du {date_debut.strftime('%d/%m')} au {date_fin.strftime('%d/%m')}"
        elements.append(Paragraph(titre_planning, planning_title_style))
        elements.append(Spacer(1, 20))

        # Planning
        data = []
        header_row = [self.grille_planning.horizontalHeaderItem(col).text() for col in range(self.grille_planning.columnCount())]
        data.append(header_row)

        for row in range(self.grille_planning.rowCount()):
            row_data = []
            for col in range(self.grille_planning.columnCount()):
                item = self.grille_planning.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        planning_table = Table(data, repeatRows=1)

        # Style du planning
        planning_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#007A33")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ]))
        elements.append(planning_table)

        # Génération du PDF
        doc.build(elements)

        # Ouverture automatique du planning dans le navigateur
        webbrowser.open(f"file://{pdf_filename}")

        print(f"PDF enregistré et ouvert : {pdf_filename}")

    def enregistrer_json(self):
        """
        Paramétrage du JSON contenant les données d'un plannings créé, exécuté après avoir cliqué "Enregistrer"
        :return:
        """
        semaine_num = self.selection_semaines.currentData()
        annee = datetime.now().year

        planning_data = {
            "semaine": f"Semaine {semaine_num}",
            "annee": annee,
            "employes": []
        }

        jours = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]

        for row in range(self.grille_planning.rowCount()):
            nom_item = self.grille_planning.item(row, 0)
            if nom_item:
                nom = nom_item.text().strip()
                horaires = {}
                for i, jour in enumerate(jours):
                    cell = self.grille_planning.item(row, i + 1)
                    horaires[jour] = cell.text() if cell else ""
                planning_data["employes"].append({
                    "nom": nom,
                    "horaires": horaires
                })

        json_dir = resource_path("Data/Plannings_json")
        os.makedirs(json_dir, exist_ok=True)

        json_filename = os.path.join(json_dir, f"planning_semaine{semaine_num}_{annee}.json")
        with open(json_filename, "w", encoding="utf-8") as f:
            json.dump(planning_data, f, indent=4, ensure_ascii=False)

        print(f"Planning JSON sauvegardé : {json_filename}")

    def get_dates_semaine(self, semaine_num):
        """Retourne la date de début (lundi) et de fin (dimanche) d'une semaine ISO donnée."""
        annee = date.today().year
        date_debut = date.fromisocalendar(annee, semaine_num, 1)  # Lundi
        date_fin = date_debut + timedelta(days=6)  # Dimanche
        return date_debut, date_fin

    def nettoyer_planning(self):
        """
        Remplacer les cellules ayant une valeur nulle ou non reconnu par une chaîne de char vide: "".
        :return:
        """
        row_count = self.grille_planning.rowCount()
        col_count = self.grille_planning.columnCount()

        for row in range(row_count):
            for col in range(1, col_count - 1):  # On saute col 0 (nom) et dernière col (Total)
                item = self.grille_planning.item(row, col)
                if item is None:
                    item = QTableWidgetItem()
                    self.grille_planning.setItem(row, col, item)
                item.setText("")

    # ---RACCOURCIS---

    # Ctrl C
    def copier_cellules_selectionnees(self):
        """
        Récupère les données dans le rectangle sélectionné et les sauvegarde dans le presse-papier.
        """
        selected_ranges = self.grille_planning.selectedRanges()
        if not selected_ranges:
            return

        copied_text = ""
        selection = selected_ranges[0]

        for row in range(selection.topRow(), selection.bottomRow() + 1):
            row_data = []
            for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                item = self.grille_planning.item(row, col)
                row_data.append(item.text() if item else "")
            copied_text += '\t'.join(row_data) + '\n'

        QGuiApplication.clipboard().setText(copied_text.strip())

    # Ctrl V
    def coller_cellules_selectionnees(self):
        clipboard_text = QGuiApplication.clipboard().text()
        lignes = clipboard_text.strip().split('\n')

        selected = self.grille_planning.selectedRanges()
        if not selected:
            return

        top_row = selected[0].topRow()
        left_col = selected[0].leftColumn()

        for i, ligne in enumerate(lignes):
            cellules = ligne.split('\t')
            for j, texte in enumerate(cellules):
                row = top_row + i
                col = left_col + j

                if row < self.grille_planning.rowCount() and col < self.grille_planning.columnCount():
                    item = self.grille_planning.item(row, col)
                    if not item:
                        item = QTableWidgetItem()
                        self.grille_planning.setItem(row, col, item)

                    texte = texte.strip()
                    key = (row, col)

                    # Historique
                    if not self.historique[key] or self.historique[key][-1] != texte:
                        self.historique[key].append(texte)
                        self.redo_stack[key] = []

                    self.grille_planning.blockSignals(True)
                    item.setText(texte)
                    self.grille_planning.blockSignals(False)

                    self.calculer_total_ligne(row)

    def effacer_cellules_selectionnees(self):
        """Réinitialise les cellules sélectionnées à une chaîne vide et conserve l'historique."""
        selected_ranges = self.grille_planning.selectedRanges()
        lignes_modifiees = set()

        for selection in selected_ranges:
            for row in range(selection.topRow(), selection.bottomRow() + 1):
                for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                    key = (row, col)
                    item = self.grille_planning.item(row, col)

                    if not item:
                        item = QTableWidgetItem("")
                        self.grille_planning.setItem(row, col, item)

                    # Sauvegarde dans l'historique AVANT de modifier
                    current_text = item.text()
                    if not self.historique[key] or current_text != self.historique[key][-1]:
                        self.historique[key].append(current_text)
                        self.redo_stack[key] = []

                    # Appliquer la suppression
                    self.grille_planning.blockSignals(True)
                    item.setText("")
                    self.grille_planning.blockSignals(False)

                    # Ajouter l'état vide à l'historique
                    self.historique[key].append("")
                    self.redo_stack[key] = []
                    lignes_modifiees.add(row)

        for row in lignes_modifiees:
            self.calculer_total_ligne(row)

        self.apply_table_style()

    # Ctrl X
    def couper_cellules_selectionnees(self):
        """Copie les cellules sélectionnées puis les efface."""
        self.copier_cellules_selectionnees()
        self.effacer_cellules_selectionnees()

    # Ctrl Z
    def retour_en_arriere(self):
        """Revenir à l'état précédent des cellules sélectionnées (Ctrl+Z)."""
        selected_ranges = self.grille_planning.selectedRanges()

        for selection in selected_ranges:
            for row in range(selection.topRow(), selection.bottomRow() + 1):
                for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                    key = (row, col)
                    if key in self.historique and len(self.historique[key]) > 1:
                        self.redo_stack[key].append(self.historique[key][-1])
                        self.historique[key].pop()
                        previous_value = self.historique[key][-1]

                        self.grille_planning.blockSignals(True)
                        self.grille_planning.item(row, col).setText(previous_value)
                        self.grille_planning.blockSignals(False)

                # Recalcul total à chaque ligne modifiée
                self.calculer_total_ligne(row)

        self.apply_table_style()

    # Ctrl Y
    def refaire(self):
        """Refaire l'action sur les cellules sélectionnées (Ctrl+Y)."""
        selected_ranges = self.grille_planning.selectedRanges()

        for selection in selected_ranges:
            for row in range(selection.topRow(), selection.bottomRow() + 1):
                for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                    key = (row, col)
                    if key in self.redo_stack and self.redo_stack[key]:
                        value_to_redo = self.redo_stack[key].pop()
                        self.historique[key].append(value_to_redo)

                        self.grille_planning.blockSignals(True)
                        self.grille_planning.item(row, col).setText(value_to_redo)
                        self.grille_planning.blockSignals(False)

                # Recalcul total ici aussi
                self.calculer_total_ligne(row)

        self.apply_table_style()

    # UI - Gestion Personnel
    def lancement_interface_employes(self):
        """
        Lance l'interface de gestion du personnel, MAJ le planning après femeture de cette interface
        """
        self.dialog = self.gestion_employes
        if self.dialog.exec() == QDialog.Accepted:  # Si on a cliqué "Sauvegarder et fermer"
            self.load_employees()  # Recharge les employés
            self.update_employes_column() #maj grille

    #---Style Sheets---

    def apply_headers_style(self):
        self.grille_planning.horizontalHeader().setStyleSheet("""
            QHeaderView::section {
                background-color: #007A33;
                color: white;
                font-family: 'Segoe UI';
                font-size: 14pt;
                padding: 4px;
            }
        """)

    def apply_table_style(self):
        self.grille_planning.setAlternatingRowColors(True)
        self.grille_planning.setStyleSheet("""
            QTableWidget {
                background-color: #F5FFF3;  /* Fond général de la table */
                alternate-background-color: #E5F7E0;  /* Fond alterné des lignes */
                gridline-color: #CCCCCC;  /* Couleur des lignes de grille */
                font-family: 'Segoe UI';
                font-size: 14pt;
                color: black;
            }

            QTableWidget::item {
                padding: 6px;
            }

            /* Style des cellules sélectionnées */
            QTableWidget::item:selected {
                background-color: transparent;  /* Couleur de fond pour les cellules sélectionnées */
                color: black;  /* Couleur du texte pour les cellules sélectionnées */
            }

            /* Style lorsque la cellule est activée par un clic */
            QTableWidget::item:selected:active {
                background-color: transparent;  /* Fond transparent pour les cellules activées */
                color: black;  /* Rendre le texte noir même si la cellule est activée */
            }

            /* Effet au survol d'une cellule (hover) */
            QTableWidget::item:hover {
                background-color: #D9EAD3;  /* Fond plus clair au survol */
                color: black;  /* Couleur du texte */
            }
        """)

    def apply_font_to_table(self):
        """
        Applique une police uniforme et un centrage à toutes les cellules de la table.
        Si une cellule est vide (aucun item), un QTableWidgetItem est créé.
        """
        font = QFont('Segoe UI', 14)

        for row in range(self.grille_planning.rowCount()):
            for col in range(self.grille_planning.columnCount()):
                item = self.grille_planning.item(row, col)

                # Créer un item vide si la cellule n'en a pas encore
                if item is None:
                    item = QTableWidgetItem()
                    self.grille_planning.setItem(row, col, item)

                item.setFont(font)
                item.setTextAlignment(Qt.AlignCenter)

    def apply_row_colors(self):
        """
        Applique une couleur de fond alternée aux lignes de la table
        (effet 'zébré' pour améliorer la lisibilité).
        """
        couleur_pair = QColor("#D9EAD3")
        couleur_impair = QColor("#B6D7A8")

        for row in range(self.grille_planning.rowCount()):
            couleur = couleur_pair if row % 2 == 0 else couleur_impair

            for col in range(self.grille_planning.columnCount()):
                item = self.grille_planning.item(row, col)
                if item is None:
                    item = QTableWidgetItem()
                    self.grille_planning.setItem(row, col, item)

                item.setBackground(QBrush(couleur))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())