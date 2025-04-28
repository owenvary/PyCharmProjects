from PySide6.QtWidgets import QApplication, QDialog, QVBoxLayout, QListWidget, QPushButton, QCheckBox, QDialogButtonBox, QLabel, QMessageBox
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

class SelectionMails(QDialog):
    def __init__(self, employes, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sélectionner les employés")

        # Liste des employés avec checkboxes
        self.selected_employees = []

        layout = QVBoxLayout(self)

        # Créez une liste d'éléments avec des cases à cocher
        self.checkbox_list = []

        # Ajouter tous les employés de la liste "employes"
        for employe in employes:
            checkbox = QCheckBox(employe['nom'])
            checkbox.setChecked(True)
            self.checkbox_list.append(checkbox)
            layout.addWidget(checkbox)

        # Ajouter Codir en plus dans la liste de sélection
        codir_checkbox = QCheckBox("Sophie")
        codir_checkbox.setChecked(True)
        self.checkbox_list.append(codir_checkbox)
        layout.addWidget(codir_checkbox)

        # Ajouter les boutons "Tout cocher" et "Tout décocher"
        self.select_all_button = QPushButton("Tout cocher")
        self.deselect_all_button = QPushButton("Tout décocher")

        # Connecter les boutons à leurs actions respectives
        self.select_all_button.clicked.connect(self.check_all)
        self.deselect_all_button.clicked.connect(self.uncheck_all)

        layout.addWidget(self.select_all_button)
        layout.addWidget(self.deselect_all_button)

        # Boutons OK et Annuler
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def get_selected_employees(self):
        # Retourner la liste des employés sélectionnés
        self.selected_employees = [checkbox.text() for checkbox in self.checkbox_list if checkbox.isChecked()]
        return self.selected_employees

    def check_all(self):
        # Cocher toutes les cases
        for checkbox in self.checkbox_list:
            checkbox.setChecked(True)

    def uncheck_all(self):
        # Décocher toutes les cases
        for checkbox in self.checkbox_list:
            checkbox.setChecked(False)