from PySide6.QtGui import QFont, QKeySequence, QShortcut, QIcon
from PySide6.QtWidgets import (
    QWidget, QTableWidget, QPushButton, QHeaderView,
    QComboBox, QVBoxLayout, QHBoxLayout, QMessageBox, QDialog, QTableWidgetItem
)
from PySide6.QtCore import Qt
import json
import os
import sys

def resource_path(relative_path):
    """Retourne le chemin absolu d’un fichier compatible dev/exe"""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)

class GestionEmployes(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.fast_planning_img = resource_path("Images/FastPlanning_logo.png")
        self.DATA_DIR = resource_path("Data/Employes_json")
        self.EMPLOYEES_FILE = resource_path("Data/Employes_json/employees.json")

        # Header + dimension fenêtre
        self.setWindowTitle("Liste des employés")
        self.setMinimumSize(800, 600)
        # Police utilisée
        self.setFont(QFont("Segoe UI", 14))
        self.setWindowIcon(QIcon(self.fast_planning_img))

        # Instance interface
        self.layout = QVBoxLayout(self)

        # Création du tableau nx2
        self.table = QTableWidget()
        self.table.setColumnCount(4)  # Ajouter une colonne pour le contrat

        # Headers tableau
        self.table.setHorizontalHeaderLabels(["Nom", "Email", "Contrat", "Déplacer"])
        self.table.horizontalHeader().setStretchLastSection(False)

        # Réglage des colonnes
        self.table.setColumnWidth(0, 150)  # Largeur de la colonne "Nom"
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)  # Email
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Contrat
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Déplacer

        # Intégration du tableau
        self.layout.addWidget(self.table)

        # Création des boutons
        self.btn_ajouter = QPushButton("Ajouter un employé")
        self.btn_supprimer = QPushButton("Supprimer l'employé sélectionné")
        self.btn_sauvegarder = QPushButton("Sauvegarder les modifications")

        # Intégration des boutons
        for btn in (self.btn_ajouter, self.btn_sauvegarder, self.btn_supprimer):
            self.layout.addWidget(btn)

        # EventListeners : btn -> fonctions
        self.btn_ajouter.clicked.connect(self.ajouter_employe)
        self.btn_supprimer.clicked.connect(self.supprimer_employe)
        self.table.itemChanged.connect(self.mettre_a_jour_donnees)  # Connecter la modification des cellules
        self.btn_sauvegarder.clicked.connect(self.sauvegarder_et_fermer)

        self.raccourci_ctrl_s = QShortcut(QKeySequence("Ctrl+S"), self)
        self.raccourci_ctrl_s.activated.connect(self.sauvegarder_et_fermer)

        # Charger les données depuis le JSON
        self.employes = []
        self.load_employes()

    def load_employes(self):
        """Charge les employés depuis EMPLOYEES_FILE, initialise une liste vide si le fichier est inexistant."""
        directory = os.path.dirname(self.EMPLOYEES_FILE)
        if not os.path.exists(directory):
            os.makedirs(directory)

        if os.path.exists(self.EMPLOYEES_FILE):
            try:
                with open(self.EMPLOYEES_FILE, "r", encoding="utf-8") as f:
                    self.employes = json.load(f)

                # Ajouter le champ "contrat" pour chaque employé si absent
                for employe in self.employes:
                    if "contrat" not in employe:
                        employe["contrat"] = "35h"  # Valeur par défaut pour le contrat

            except Exception as e:
                QMessageBox.warning(self, "Erreur", f"Impossible de charger les employés : {e}")
                self.employes = []
        else:
            self.employes = []

        # Mise à jour du tableau
        self.update_table()

    def update_table(self):
        """Remplit le tableau avec la liste actuelle d'employés et ajoute les flèches pour chaque ligne."""
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.employes))

        for row, employe in enumerate(self.employes):
            self.table.setItem(row, 0, QTableWidgetItem(employe.get("nom", "")))
            self.table.setItem(row, 1, QTableWidgetItem(employe.get("email", "")))

            # Création du menu déroulant (QComboBox) pour le contrat
            contrat_combobox = QComboBox()
            contrat_combobox.addItems(["33h", "35h", "39h", "ALTERNANT", "ÉTUDIANT", "PATRON"])
            contrat_combobox.setCurrentText(employe.get("contrat", "35h"))  # Sélectionner la valeur actuelle

            contrat_combobox.currentTextChanged.connect(lambda text, row=row: self.mettre_a_jour_contrat(text, row))

            self.table.setCellWidget(row, 2, contrat_combobox)

            # Créer le layout pour les deux flèches
            cell_widget = QWidget()
            layout = QHBoxLayout()
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(2)

            btn_up = QPushButton("\u25B2")  # 🔼
            btn_down = QPushButton("\u25BC")  # 🔽

            btn_up.setFixedSize(30, 25)
            btn_down.setFixedSize(30, 25)

            # Styles des btn up & down
            btn_up.setStyleSheet("padding: 0px; margin: 0px;")
            btn_down.setStyleSheet("padding: 0px; margin: 0px;")

            btn_up.clicked.connect(lambda _, r=row: self.deplacer_haut(r))
            btn_down.clicked.connect(lambda _, r=row: self.deplacer_bas(r))

            layout.addWidget(btn_up)
            layout.addWidget(btn_down)

            layout.setAlignment(Qt.AlignCenter)
            cell_widget.setLayout(layout)

            self.table.setCellWidget(row, 3, cell_widget)

        self.table.blockSignals(False)

    def mettre_a_jour_contrat(self, text, row):
        """Met à jour le contrat de l'employé dans la liste interne."""
        if row < len(self.employes):
            self.employes[row]["contrat"] = text

    def mettre_a_jour_donnees(self, item):
        """Met à jour les données de nom et email lors de la modification dans le tableau."""
        row = item.row()
        column = item.column()

        if column == 0:  # Nom
            self.employes[row]["nom"] = item.text()
        elif column == 1:  # Email
            self.employes[row]["email"] = item.text()

    def deplacer_haut(self, row):
        """Déplace la ligne sélectionnée vers le haut."""
        if row > 0:
            self.employes[row], self.employes[row - 1] = self.employes[row - 1], self.employes[row]
            self.update_table()

    def deplacer_bas(self, row):
        """Déplace la ligne sélectionnée vers le bas."""
        if row < len(self.employes) - 1:
            self.employes[row], self.employes[row + 1] = self.employes[row + 1], self.employes[row]
            self.update_table()

    def save_employes(self):
        """Sauvegarde les employés dans le fichier JSON."""
        try:
            with open(self.EMPLOYEES_FILE, "w", encoding="utf-8") as f:
                json.dump(self.employes, f, ensure_ascii=False, indent=4)
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Erreur lors de la sauvegarde : {e}")

    def ajouter_employe(self):
        """Ajoute un nouvel employé par défaut au tableau."""
        self.employes.append({
            "nom": "Nouvel Employé",
            "email": "nouveau@mail.com",
            "contrat": "35h"  # Valeur par défaut pour le contrat
        })
        self.update_table()

    def supprimer_employe(self):
        """Supprime l'employé sélectionné après confirmation du pop-up."""
        row = self.table.currentRow()
        if row >= 0:
            confirm = QMessageBox.question(
                self, "Confirmation", "Supprimer cet employé ?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if confirm == QMessageBox.Yes:
                self.employes.pop(row)
                self.update_table()
        else:
            QMessageBox.warning(self, "Alerte", "Aucun employé sélectionné.")

    def sauvegarder_et_fermer(self):
        """Sauvegarde les employés puis ferme la fenêtre."""
        self.save_employes()
        QMessageBox.information(self, "Succès", "Employés enregistrés avec succès.")
        self.accept()
