# PlanningCarrefour

## Description

**PlanningCarrefour** est une application Python permettant la génération de plannings tout en prenant compte le turn-over. Elle inclut également l'envoi d'emails pour notifier les employés de leurs horaires. Ce projet vise à faciliter la gestion des plannings dans un environnement professionnel.

## Fonctionnalités

- **Gestion des employés** : Ajout, modification et suppression des employés. Incluant la gestion de leurs emails pour l'envoi.
- **Génération des plannings** : Création automatique des plannings hebdomadaires avec une option de chargement d'anciens plannings pour faciliter la création.
- **Envoi d'emails** : Envoi d'emails aux employés pour les informer de leurs horaires.
- **Interface simple** : Facilité d'utilisation pour les utilisateurs non techniques en incluant des raccourcis tel que copier, colelr, undo, redo.

## Structure du projet

```text
pythonProject/
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
├── logo_carrefour_city.png   # Logo du projet
├── main.py                   # Script principal du projet
└── test.py                   # Script pour les tests
```



