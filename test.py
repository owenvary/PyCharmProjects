import sys
import json
import os
import locale
import re
import webbrowser
from collections import defaultdict

from PySide6.QtGui import QFont, QBrush, QColor, QKeyEvent, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTableWidget, QLabel, QPushButton,
    QGridLayout, QTableWidgetItem, QHeaderView, QSizePolicy,
    QComboBox, QVBoxLayout, QHBoxLayout, QMessageBox, QDialog, QAbstractItemView,
)

from PySide6.QtCore import Qt, QEvent
from datetime import datetime, timedelta, date

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

from gestion_employes import GestionEmployes
from envoi_mails import EnvoiPlanning



class Window(QMainWindow):
    def __init__(self):
        super().__init__()

        # Configuration générale
        locale.setlocale(locale.LC_TIME, 'fr_FR')  # Dates en français

        self.gestion_employes = GestionEmployes()

        self.BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        self.PLANNINGS_JSON_DIR = os.path.join(self.BASE_DIR, "Data", "Plannings_JSON")
        self.PLANNINGS_PDF_DIR = os.path.join(self.BASE_DIR, "Data", "Plannings_PDF")
        os.makedirs(self.PLANNINGS_JSON_DIR, exist_ok=True)
        os.makedirs(self.PLANNINGS_PDF_DIR, exist_ok=True)

        self.load_employees()
        self.nb_employees = len(self.employees)

        self.setWindowTitle("Planning")
        self.setWindowState(Qt.WindowMaximized)

        # Panneau central
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout(self.central_widget)

        # Panneau du haut (sélection semaine + boutons charger, employés)
        top_layout = QHBoxLayout()

        # Menu déroulant de sélection de la semaine
        self.selection_semaines = QComboBox()
        self.selection_semaines.setMinimumWidth(350)
        self.integrate_week_selections(date.today().year)

        # Bouton charger planning
        self.btn_charger = QPushButton("Charger planning")
        self.btn_charger.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        # Bouton employés
        self.btn_employes = QPushButton("Liste des employés")
        self.btn_employes.clicked.connect(self.lancement_interface_employes)

        # Layout pour menu déroulant + bouton charger (partie gauche)
        left_layout = QHBoxLayout()
        left_layout.addWidget(self.selection_semaines)
        left_layout.addSpacing(15)
        left_layout.addWidget(self.btn_charger)

        # Ajout des éléments au top layout
        top_layout.addLayout(left_layout)
        top_layout.addStretch()
        top_layout.addWidget(self.btn_employes) # (partie droite)

        # Centre : Titre + Grille de planning
        center_layout = QGridLayout()

        self.semaine_selected = self.selection_semaines.currentText()
        self.selection_semaines.currentIndexChanged.connect(self.semaine_choisie)

        self.label_title = QLabel(f"Planning - {self.semaine_selected}")
        self.label_title.setFont(QFont('Segoe UI', 16))
        center_layout.addWidget(self.label_title, 0, 0, alignment=Qt.AlignCenter)

        # Grille principale
        self.table = QTableWidget(self.nb_employees, 9)
        self.maj_headers_planning()

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Sélection multiple des cellules de la grille planning
        self.table.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.installEventFilter(self)

        center_layout.addWidget(self.table, 1, 0)

        # Apparence du planning
        self.apply_headers_style()
        self.apply_table_style()
        self.apply_font_to_table()
        self.apply_row_colors()

        # Bas : Boutons enregistrer et envoyer
        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(40)

        self.btn_sauvegarder = QPushButton("Enregistrer planning")
        self.btn_envoyer = QPushButton("Envoyer le planning")

        for btn in (self.selection_semaines, self.btn_charger, self.btn_employes, self.btn_sauvegarder, self.btn_envoyer):
            btn.setFont(QFont('Segoe UI', 14))
            btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        bottom_layout.addStretch()
        bottom_layout.addWidget(self.btn_sauvegarder)
        bottom_layout.addWidget(self.btn_envoyer)

        # Assemblage final
        main_layout.addLayout(top_layout)
        main_layout.addLayout(center_layout)
        main_layout.addLayout(bottom_layout)

        # Initialisation de la colonne 'Employé'
        self.update_employes_column()

        # Création des EventListeners
        self.table.cellChanged.connect(self.eventListener_chgt_cellule)
        self.btn_charger.clicked.connect(self.load_planning)
        self.btn_sauvegarder.clicked.connect(self.enregistrer_planning)

        self.btn_envoyer.clicked.connect(self.envoi_planning)
        # Historique pour les valeurs des cellules
        self.historique = defaultdict(list)
        self.redo_stack = defaultdict(list)  # Stack pour refaire les actions (Ctrl+Y)

        # Initialisation de l'historique
        self.init_historique()

        # Raccourcis Ctrl+Z et Ctrl+Y
        self.raccourci_ctrl_z = QShortcut(QKeySequence("Ctrl+Z"), self)
        self.raccourci_ctrl_z.activated.connect(self.retour_en_arriere)

        self.raccourci_ctrl_y = QShortcut(QKeySequence("Ctrl+Y"), self)
        self.raccourci_ctrl_y.activated.connect(self.refaire)

    def init_historique(self):
        """Initialise l'historique pour toutes les cellules de la table."""
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if not item:
                    item = QTableWidgetItem("")  # Initialiser la cellule si elle est vide
                    self.table.setItem(row, col, item)
                # Initialiser l'historique avec le texte de la cellule au début
                self.historique[(row, col)] = [item.text()]
                self.redo_stack[(row, col)] = []  # Initialisation du stack de redo

    def keyPressEvent(self, event):
        current_item = self.table.currentItem()
        if current_item:
            row = self.table.currentRow()
            col = self.table.currentColumn()
            texte_courant = current_item.text()

            historique_cell = self.historique[(row, col)]

            # Éviter les doublons consécutifs
            if not historique_cell or texte_courant != historique_cell[-1]:
                historique_cell.append(texte_courant)
                self.redo_stack[(row, col)] = []

        super().keyPressEvent(event)

    def retour_en_arriere(self):
        """Revenir à l'état précédent des cellules sélectionnées (Ctrl+Z)."""
        selected_ranges = self.table.selectedRanges()

        for selection in selected_ranges:
            for row in range(selection.topRow(), selection.bottomRow() + 1):
                for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                    key = (row, col)
                    if key in self.historique and len(self.historique[key]) > 1:
                        self.redo_stack[key].append(self.historique[key][-1])
                        self.historique[key].pop()
                        previous_value = self.historique[key][-1]

                        self.table.blockSignals(True)
                        self.table.item(row, col).setText(previous_value)
                        self.table.blockSignals(False)

    def refaire(self):
        """Refaire l'action sur les cellules sélectionnées (Ctrl+Y)."""
        selected_ranges = self.table.selectedRanges()

        for selection in selected_ranges:
            for row in range(selection.topRow(), selection.bottomRow() + 1):
                for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                    key = (row, col)
                    if key in self.redo_stack and self.redo_stack[key]:
                        value_to_redo = self.redo_stack[key].pop()
                        self.historique[key].append(value_to_redo)

                        self.table.blockSignals(True)
                        self.table.item(row, col).setText(value_to_redo)
                        self.table.blockSignals(False)

    def envoi_planning(self):
        # Demander confirmation avant d'envoyer l'email
        confirm = QMessageBox.question(
            self, "Confirmation", "Voulez-vous envoyer le planning ?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if confirm == QMessageBox.Yes:
            # Si l'utilisateur choisit "Oui", envoyer l'email
            envoi_planning = EnvoiPlanning(self)
            gestion_employes = GestionEmployes
            success = envoi_planning.send_email_with_pdf()

            # Affichage du pop-up de confirmation
            if success:
                QMessageBox.information(self, "Succès", "Le planning a été envoyé avec succès")
            else:
                QMessageBox.information(self, "Erreur", "Une erreur est survenue lors de l'envoi du planning.")
        else:
            # Si l'utilisateur choisit "Non", on annule l'envoi
            QMessageBox.information(self, "Annulé", "L'envoi du planning a été annulé.")

    def calculer_total_ligne(self, row):
        total = 0.0
        total_col = self.table.columnCount() - 1  # Colonne pour le total

        # Récupérer le nom de l'employé pour cette ligne (1ère colonne = colonne 0)
        nom_item = self.table.item(row, 0)
        nom_employe = nom_item.text().strip() if nom_item else ""

        for col in range(total_col):  # On exclut la dernière colonne
            item = self.table.item(row, col)
            if item:
                texte = item.text().strip()
                if texte:
                    try:
                        if texte in ("CP", "CGP"):
                            # Cas congé : condition spéciale pour Jean-Claude
                            if nom_employe == "Jean-Claude":
                                total += 33 / 6
                            else:
                                total += 35 / 6
                        elif texte in ("AFORMANCE", "CFA", "COURS"):
                            total += 7
                        elif isinstance(texte, str):
                            # Traitement normal des horaires
                            parties = re.split(r"[-\s]+", texte)
                            heures = list(map(float, parties))
                            for i in range(0, len(heures) - 1, 2):
                                debut = heures[i]
                                fin = heures[i + 1]
                                total += abs(fin - debut)
                        else:
                            print(f"Texte invalide dans la cellule ({row}, {col}): {texte}")
                    except ValueError:
                        pass  # Ignore les saisies non valides

        # Met à jour ou crée la cellule de total
        total_item = self.table.item(row, total_col)
        if total_item is None:
            total_item = QTableWidgetItem(f"{round(total, 2)}")
            total_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(row, total_col, total_item)
        else:
            total_item.setText(f"{round(total, 2)}")

    def eventListener_chgt_cellule(self, row, col):
        """Transforme les espaces en ' - ' selon la logique définie, sauf si un tiret est déjà présent.
        Enregistre uniquement le texte final transformé dans l'historique.
        """
        item = self.table.item(row, col)
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
                self.table.blockSignals(True)
                item.setText(result_text)
                self.table.blockSignals(False)

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

    def update_employes_column(self):
        noms = [emp["nom"] for emp in self.employees]

        self.table.setRowCount(len(noms))

        for row in range(len(noms)):
            item = self.table.item(row, 0)
            if item is None:
                item = QTableWidgetItem()
                self.table.setItem(row, 0, item)

            item.setText(noms[row])
            item.setFont(QFont('Segoe UI', 14))
            item.setTextAlignment(Qt.AlignCenter)

            # Rendre la cellule non éditable afin de conserver ses propriétés
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)

        # Re-appliquer le design des cellules ayant été modifié par les selections
        self.apply_row_colors()
        self.apply_font_to_table()

    def load_employees(self):
        employees_file = os.path.join(self.BASE_DIR, "Data", "Employes_json", "employees.json")
        if os.path.exists(employees_file):
            with open(employees_file, "r", encoding="utf-8") as f:
                self.employees = json.load(f)
        else:
            self.employees = []
            self.update_employes_column()

    def lancement_interface_employes(self):
        self.dialog = self.gestion_employes
        if self.dialog.exec() == QDialog.Accepted:  # Si on a cliqué "Sauvegarder"
            self.load_employees()  # Recharge les employés
            self.update_employes_column() #maj grille

    def semaine_choisie(self, index):
        """ met à jour le planning et le titre selon la semaine choisie
        :param index: numéro de la semaine choisie"""
        self.semaine_selected = self.selection_semaines.itemText(index)
        self.set_planning_title()
        self.maj_headers_planning()

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
        self.table.setHorizontalHeaderLabels(headers)

    def get_start_of_week(self):
        """
        Retourne le premier jour (lundi) de la semaine sélectionnée dans le menu déroulant.
        Si aucune semaine n'est sélectionnée, retourne la date du jour.
        """
        if self.selection_semaines.currentData() is None:
            return date.today()

        week_number = int(self.selection_semaines.currentData())  # Numéro de semaine (1 à 52)
        return date.fromisocalendar(date.today().year, week_number, 1)  # 1 = Lundi

    def eventFilter(self, source, event): # Méthode de PyQt
        """
        Intercepte les événements clavier sur la table pour gérer le copier-coller personnalisé.
        Si une combinaison Ctrl+C ou Ctrl+V est détectée, appelle la méthode correspondante.
        """
        if source == self.table and isinstance(event, QKeyEvent):
            # Gestion du copier (Ctrl + C)
            if event.matches(QKeySequence.Copy):
                self.copier_cellules_selectionnees()
                return True  # Empêche le comportement par défaut

            # Gestion du coller (Ctrl + V)
            elif event.matches(QKeySequence.Paste):
                self.table.blockSignals(True) # Pour pas que eventListeners_chgt_cell se déclenche
                self.coller_cellules_selectionnees()
                self.table.blockSignals(False)
                return True

        # Sinon, comportement normal
        return super().eventFilter(source, event)

    def copier_cellules_selectionnees(self):
        """
        Copie le contenu des cellules actuellement sélectionnées dans le planning,
        en format texte tabulé, prêt à être collé.
        """
        selection = self.table.selectedRanges()
        if not selection:
            return  # Rien n'est sélectionné

        copied_text = ""
        # On prend la première plage de sélection uniquement
        for row in range(selection[0].topRow(), selection[0].bottomRow() + 1):
            row_data = []
            for col in range(selection[0].leftColumn(), selection[0].rightColumn() + 1):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")  # On récupère le texte ou une chaîne vide
            copied_text += "\t".join(row_data) + "\n"  # Séparateur tab entre colonnes, retour ligne entre lignes

        QApplication.clipboard().setText(copied_text.strip())  # On copie dans le presse-papier système

    def coller_cellules_selectionnees(self):
        """
        Colle le contenu du presse-papier dans la table à partir de la cellule actuellement sélectionnée.
        Le format attendu est tabulé (comme ce que produit la méthode `copier_cellules_selectionnees`).
        """
        texte = QApplication.clipboard().text()
        if not texte:
            return  # Rien à coller

        lignes = texte.split("\n")
        start_row = self.table.currentRow()
        start_col = self.table.currentColumn()

        # Parcours des lignes et colonnes du texte copié
        for i, ligne in enumerate(lignes):
            colonnes = ligne.split("\t")
            for j, valeur in enumerate(colonnes):
                row = start_row + i
                col = start_col + j
                if row < self.table.rowCount() and col < self.table.columnCount():
                    item = self.table.item(row, col)
                    if item is None:
                        # Si aucune cellule n'existe à cet endroit, on en crée une
                        item = QTableWidgetItem(valeur)
                        self.table.setItem(row, col, item)
                    else:
                        # Sinon, on met simplement à jour le texte
                        item.setText(valeur)

    def integrate_week_selections(self, year):
        today = date.today()
        week = 1
        current_week = today.isocalendar()[1]  # Récupère la semaine actuelle

        while True:
            try:
                monday = date.fromisocalendar(year, week, 1)
                sunday = monday + timedelta(days=6)

                if monday < today and (today - monday).days <= 90:
                    item_text = f"Semaine {week} - du {monday.strftime('%d/%m')} au {sunday.strftime('%d/%m')} - {year}"
                    self.selection_semaines.addItem(item_text, userData=week)

                    # Si c'est la semaine actuelle, on la sélectionne
                    if week == current_week:
                        self.selection_semaines.setCurrentIndex(self.selection_semaines.count() - 1)

                elif monday >= today and (monday - today).days <= 30:
                    item_text = f"Semaine {week} - du {monday.strftime('%d/%m')} au {sunday.strftime('%d/%m')} - {year}"
                    self.selection_semaines.addItem(item_text, userData=week)

                if monday > today + timedelta(days=30):
                    break
                week += 1
            except ValueError:
                break

    def apply_headers_style(self):
        self.table.horizontalHeader().setStyleSheet("""  
            QHeaderView::section {
                background-color: #007A33;
                color: white;
                font-family: 'Segoe UI';
                font-size: 14pt;
                padding: 4px;
            }
        """)

    def apply_table_style(self):
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""  
            QTableWidget {
                background-color: #F5FFF3;
                alternate-background-color: #E5F7E0;
                gridline-color: #CCCCCC;
                font-family: 'Segoe UI';
                font-size: 14pt;
            }
            QTableWidget::item {
                padding: 6px;
            }
        """)

    def apply_font_to_table(self):
        """
        Applique une police uniforme et un centrage à toutes les cellules de la table.
        Si une cellule est vide (aucun item), un QTableWidgetItem est créé.
        """
        font = QFont('Segoe UI', 14)

        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)

                # Créer un item vide si la cellule n'en a pas encore
                if item is None:
                    item = QTableWidgetItem()
                    self.table.setItem(row, col, item)

                item.setFont(font)
                item.setTextAlignment(Qt.AlignCenter)

    def apply_row_colors(self):
        """
        Applique une couleur de fond alternée aux lignes de la table
        (effet 'zébré' pour améliorer la lisibilité).
        """
        couleur_pair = QColor("#D9EAD3")
        couleur_impair = QColor("#B6D7A8")

        for row in range(self.table.rowCount()):
            couleur = couleur_pair if row % 2 == 0 else couleur_impair

            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item is None:
                    item = QTableWidgetItem()
                    self.table.setItem(row, col, item)

                item.setBackground(QBrush(couleur))

    def load_planning(self):
        """Charge un planning depuis un fichier JSON, en fonction de la semaine sélectionnée."""

        # Récupérer la semaine sélectionnée et l'année courante
        semaine_num = self.selection_semaines.currentData()
        annee = datetime.now().year

        # Chemin du fichier JSON à charger
        filename = os.path.join(self.BASE_DIR, "Data", "Plannings_json", f"planning_semaine{semaine_num}_{annee}.json")

        # Vérifier que le fichier existe
        if not os.path.exists(filename):
            QMessageBox.warning(self, "Fichier introuvable", f"Aucun planning trouvé pour la semaine {semaine_num}.")
            return

        # Charger les données depuis le fichier
        with open(filename, "r", encoding="utf-8") as f:
            planning_data = json.load(f)

        # Nettoyer toutes les anciennes données (hors colonne Nom et Total)
        for row in range(self.table.rowCount()):
            for col in range(1, self.table.columnCount() - 1):
                self.table.setItem(row, col, QTableWidgetItem(""))

        # Liste des jours de la semaine (alignée avec les colonnes 1 à 7)
        jours = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]

        # Création d’un dictionnaire pour retrouver rapidement la ligne d’un employé à partir de son nom
        nom_to_row = {}
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item:
                nom_to_row[item.text().strip().lower()] = row

        # Répartir les données horaires dans les bonnes cellules du tableau
        for emp_data in planning_data.get("employes", []):
            nom = emp_data.get("nom", "").strip().lower()
            if nom in nom_to_row:
                row = nom_to_row[nom]
                horaires = emp_data.get("horaires", {})
                for col, jour in enumerate(jours):
                    texte = horaires.get(jour, "")
                    self.table.setItem(row, col + 1, QTableWidgetItem(texte))

        # Rafraîchir les styles après chargement
        self.apply_font_to_table()
        self.apply_row_colors()

        print(f"Planning de la semaine {semaine_num} chargé avec succès.")

    def get_dates_semaine(self, semaine_num):
        """Retourne la date de début (lundi) et de fin (dimanche) d'une semaine ISO donnée."""
        annee = date.today().year
        date_debut = date.fromisocalendar(annee, semaine_num, 1)  # Lundi
        date_fin = date_debut + timedelta(days=6)  # Dimanche
        return date_debut, date_fin

    def nettoyer_planning(self):
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()

        for row in range(row_count):
            for col in range(1, col_count - 1):  # On saute col 0 (nom) et dernière col (Total)
                item = self.table.item(row, col)
                if item is None:
                    item = QTableWidgetItem()
                    self.table.setItem(row, col, item)
                item.setText("")

    def enregistrer_planning(self):
        self.enregistrer_pdf()  # Sauvegarde PDF pour l'envoi et l'impression
        self.save_json_planning()  # Sauvegarde JSON pour le chargement

    def enregistrer_pdf(self):
        """
        Enregistre et génère le planning.
        :return:
        """
        semaine_num = self.selection_semaines.currentData()
        annee = datetime.now().year
        date_debut, date_fin = self.get_dates_semaine(semaine_num)  # Tu dois avoir une fonction ou méthode pour ça
        pdf_filename = os.path.join(self.PLANNINGS_PDF_DIR, f"planning_semaine{semaine_num}_{annee}.pdf")

        # Paramètre de la page
        doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(A4), leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
        elements = []

        # HEADER: Logo Carrefour + Titre
        logo_path = os.path.join(self.BASE_DIR, "logo_carrefour_city.png")

        try:
            logo = Image(logo_path, width=80, height=80)
        except:
            logo = Paragraph("Logo manquant", getSampleStyleSheet()["Normal"])

        # Enregistrer une nouvelle police
        try:
            pdfmetrics.registerFont(TTFont("Montserrat-Bold", os.path.join(self.BASE_DIR, "Montserrat-Bold.ttf")))
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
        header_row = [self.table.horizontalHeaderItem(col).text() for col in range(self.table.columnCount())]
        data.append(header_row)

        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
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

    def save_json_planning(self):
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

        for row in range(self.table.rowCount()):
            nom_item = self.table.item(row, 0)
            if nom_item:
                nom = nom_item.text().strip()
                horaires = {}
                for i, jour in enumerate(jours):
                    cell = self.table.item(row, i + 1)
                    horaires[jour] = cell.text() if cell else ""
                planning_data["employes"].append({
                    "nom": nom,
                    "horaires": horaires
                })

        json_dir = os.path.join(self.BASE_DIR, "Data", "Plannings_json")
        os.makedirs(json_dir, exist_ok=True)

        json_filename = os.path.join(json_dir, f"planning_semaine{semaine_num}_{annee}.json")
        with open(json_filename, "w", encoding="utf-8") as f:
            json.dump(planning_data, f, indent=4, ensure_ascii=False)

        print(f"Planning JSON sauvegardé : {json_filename}")




if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    gestion_equipe = GestionEmployes()
    window.show()
    sys.exit(app.exec())