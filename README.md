# PlanningCarrefour

## Description

**PlanningCarrefour** est une application Python permettant la génération de plannings tout en prenant compte le turn-over. Elle inclut également l'envoi d'emails pour notifier les employés de leurs horaires. Ce projet vise à faciliter la gestion des plannings dans un environnement professionnel.

## Fonctionnalités

- **Gestion des employés** : Ajout, modification et suppression des employés. Incluant la gestion de leurs emails pour l'envoi ainsi que la gestiond e leur contrat.
- **Génération des plannings** : Création automatique des plannings hebdomadaires avec une option de chargement d'anciens plannings pour faciliter la création.
- **Envoi d'emails** : Envoi d'emails aux employés pour les informer de leurs horaires.
- **Interface simple** : Facilité d'utilisation pour les utilisateurs non techniques en incluant des raccourcis tel que enregistrer, copier, coller, couper, undo, redo.
- **Auto-Analyse du planning** : Analyse du contenu du planning en cours de création. Relève les sous effectifs, la confirmité des créneaux de chaque employés (total par rapport au contrat, chevauchements).
- **Infos** : Ouverture du guide d'utilisation dans le navigateur par défault sous condition de cliquer sur le bouton 'infos'

## Structure du projet

```text
pythonProject/
│
├── Images/
│   ├── FastPlanning.png      # Logo de l'application, généré par une IA
│   └── logo_carrefour_city   # Logo officiel Carrefour City
│
├── Data/
│   ├── Employes_json/        # Contient employees.json avec les données des employés
│   ├── Plannings_json/       # Fichiers JSON des plannings (planning_semaineX_202X.json) = sauvegardes du planning
│   └── Plannings_pdf/        # Fichiers PDF des plannings (planning_semaineX_202X.pdf)
│
├── mesProjets/               # Dossier contenant des fichiers Python et des configurations
│   ├── Lib/
│   ├── Scripts/
│   └── pyvenv.cfg            # Configuration de l'environnement virtuel
│
├── envoi_mails.py            # Script Python pour envoyer les mails
├── gestion_employes.py       # Script Python pour la gestion des employés
├── selection_mails.py        #Script Python pour sélectionner à qui envoyer le planning
├── main.py                   # Script principal du projet
├── test.py                   # Script pour les tests
└── pré-Guide d'utilisation FastPlanning.pdf # Guide d'utilisation step by step pdf
```



