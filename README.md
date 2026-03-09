# Générateur de Rapport Excel KPI 3G Automatisé

## Description
Ce projet est un automate local conçu pour générer des rapports KPI 3G au format Excel en se basant sur un fichier de données brutes ("raw data") et en reproduisant la structure logique d'un template donné.

**Confidentialité garantie :** Cette application est 100% locale. Elle ne fait aucun appel à des API externes, ni à des services cloud. Vos données KPI restent strictement sur votre machine.

## Fonctionnalités Principales
- **Extraction des 10 dernières heures :** L'outil scanne automatiquement la colonne *Begin Time* de vos données brutes pour extraire les 10 dernières heures d'activité enregistrées.
- **Correspondance Automatique des KPI :** Les colonnes du fichier "raw data" sont mappées automatiquement aux KPI attendus dans le Template.
- **Formatage Conditionnel Local :** L'outil recrée la logique de couleurs spécifiée (en utilisant la bibliothèque `openpyxl`) :
  - `Cell Availability (%)` >= 98.5 est **Vert**, sinon **Rouge pâle**.
  - `CSSR (CS et PS)` < 98.5 est **Rouge pâle**, sinon **Vert**.
  - `Call Drop (CS et PS)` <= 0.7 est **Vert**, sinon **Rouge pâle**.
  - `CS Voice Traffic` reste **Blanc** (sans remplissage).
- **Interface Graphique :** Une interface utilisateur claire et simple, basée sur `tkinter`, permet de sélectionner les fichiers facilement.

## Prérequis
- **Python 3.8+** doit être installé sur la machine.
- Ce projet est conçu pour être utilisé depuis VS Code (ou via la ligne de commande/explorateur Windows).

## Installation

Ouvrez un terminal dans le dossier du projet et installez les dépendances :
```bash
pip install -r requirements.txt
```

## Utilisation

### Sous Windows
Double-cliquez simplement sur le fichier `run.bat`. Ce script installera les dépendances si nécessaires et lancera l'application graphique.

### Via Ligne de Commande (ou VS Code)
Exécutez la commande suivante à la racine du projet :
```bash
python main.py
```

### Fonctionnement de l'interface
1. **Fichier Template :** Cliquez sur Parcourir et sélectionnez votre fichier modèle (ex: `Event_Performance Management-...xlsx`).
2. **Données Brutes :** Sélectionnez le fichier Excel contenant les nouvelles données brutes (ex: `Performance Management-...xlsx`).
3. **Dossier Sortie :** Indiquez le dossier où le nouveau fichier sera sauvegardé ainsi que son nom de fichier souhaité.
4. Cliquez sur le bouton **Générer le Rapport**. L'outil va traiter le fichier, générer le fichier Excel final (contenant uniquement un seul *Sheet*) et appliquer toutes les couleurs requises.

## Architecture des Fichiers
- `main.py` : Le point d'entrée du programme et l'interface utilisateur.
- `data_processor.py` : Lit et filtre les 10 dernières heures de données brutes.
- `kpi_matcher.py` : Organise et fait la correspondance entre les données brutes et les KPIs du template.
- `report_generator.py` : Utilise `openpyxl` pour générer le fichier Excel de sortie avec les styles et le formatage conditionnel.
- `run.bat` : Script de lancement rapide pour environnement Windows.
- `requirements.txt` : Liste des dépendances (`pandas`, `openpyxl`).
