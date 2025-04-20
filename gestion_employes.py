
import json
import os


from PySide6.QtGui import QFont
from PySide6.QtWidgets import ( QTableWidget, QPushButton,
    QTableWidgetItem, QHeaderView, QVBoxLayout, QMessageBox, QDialog
)



class GestionEmployes(QDialog):



    def __init__(self, parent=None):
        super().__init__(parent)

        # Header + dimension fenêtre
        self.setWindowTitle("Liste des employés")
        self.setMinimumSize(800, 600)

        self.BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # Répertoire de base
        self.DATA_DIR = os.path.join(self.BASE_DIR, "Data", "Employes_json")  # Dossier des employés
        self.EMPLOYEES_FILE = os.path.join(self.DATA_DIR, "employees.json")  # Chemin complet du fichier JSON

        # Police utilisée
        self.setFont(QFont("Segoe UI", 12))

        # Instance interface
        self.layout = QVBoxLayout(self)

        # Création du tableau nx2
        self.table = QTableWidget()
        self.table.setColumnCount(2)

        # Headers tableau
        self.table.setHorizontalHeaderLabels(["Nom", "Email"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

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
        self.table.itemChanged.connect(self.mettre_a_jour_donnees)
        self.btn_sauvegarder.clicked.connect(self.sauvegarder_et_fermer)

        # Charger les données depuis le JSON
        self.employes = []
        self.load_employes()

    def sauvegarder_et_fermer(self):
        """ Assure la fermeture et déclenche la sauvegarde des données employés après que btn_sauvegatder ait été cliqué."""
        self.save_employes()
        QMessageBox.information(self, "Succès", "Employés enregistrés avec succès.")
        self.accept()

    def load_employes(self):
        """Charge les employés depuis EMPLOYEES_FILE, initialise une liste vide si le fichier est inexistant."""
        # Vérifier si le dossier existe, sinon le créer
        directory = os.path.dirname(self.EMPLOYEES_FILE)
        if not os.path.exists(directory):
            os.makedirs(directory)

        # Vérifier si le fichier existe puis l'ouvrir en mode lecture
        if os.path.exists(self.EMPLOYEES_FILE):
            try:
                with open(self.EMPLOYEES_FILE, "r", encoding="utf-8") as f:
                    self.employes = json.load(f)
            # Génère un pop-up si le chargement à échouer
            except Exception as e:
                QMessageBox.warning(self, "Erreur", f"Impossible de charger les employés : {e}")
                self.employes = []
        else:
            # Si le fichier n'existe pas, on l'initialise avec une liste vide
            self.employes = []

        # Mise à jour du tableau
        self.update_table()

    def update_table(self):
        """Remplit le tableau avec la liste actuelle d'employés."""
        self.table.blockSignals(True)  # Éviter qu'une fonction se déclenche due à un changement dans une cellule
        self.table.setRowCount(len(self.employes))
        for row, employe in enumerate(self.employes):
            self.table.setItem(row, 0, QTableWidgetItem(employe.get("nom", ""))) # Chaîne vide par défault
            self.table.setItem(row, 1, QTableWidgetItem(employe.get("email", "")))
        self.table.blockSignals(False) # Réactiver les signaux

    def save_employes(self):
        """Sauvegarde les employés dans le fichier JSON."""
        try:
            with open(self.EMPLOYEES_FILE, "w", encoding="utf-8") as f:
                # ensure_ascii = False afin d'éviter les problèmes avec les accents
                json.dump(self.employes, f, ensure_ascii=False, indent=4)
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Erreur lors de la sauvegarde : {e}")

    def get_noms_employes(self):
        return [employe["nom"] for employe in self.employes]

    def ajouter_employe(self):
        """Ajoute un nouvel employé par défaut au tableau. Création, mise à jour puis sauvegarde"""
        self.employes.append({"nom": "Nouvel Employé", "email": "nouveau@mail.com"})
        self.update_table()
        self.save_employes()

    def supprimer_employe(self):
        """Supprime l'employé sélectionné après confirmation du pop-up"""
        row = self.table.currentRow()
        if row >= 0:
            confirm = QMessageBox.question(
                self, "Confirmation", "Supprimer cet employé ?",
                QMessageBox.Yes | QMessageBox.No
            )
            if confirm == QMessageBox.Yes:
                self.employes.pop(row)
                self.update_table()
                self.save_employes()
        else:
            QMessageBox.warning(self, "Alerte", "Aucun employé sélectionné.")

    def mettre_a_jour_donnees(self, item):
        """Met à jour les données internes quand une cellule du tableau est modifiée."""
        row = item.row()
        col = item.column()

        if row < len(self.employes):
            if col == 0:
                self.employes[row]["nom"] = item.text()
            elif col == 1:
                self.employes[row]["email"] = item.text()

            # Enregistre les modifications dans le fichier JSON
            with open(self.EMPLOYEES_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.employes, f, indent=4, ensure_ascii=False)

    def charger_employes(self):
        """
        Met à jour le tableau noms - mails des employés après avoir chargé els données depuis le JSON
        """
        path = os.path.join(self.BASE_DIR, "Data", "employes.json")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                employes = json.load(f)

            self.table.setRowCount(0)
            for nom in employes:
                row_position = self.table.rowCount()
                self.table.insertRow(row_position)
                self.table.setItem(row_position, 0, QTableWidgetItem(nom))
            self.update_table()
