# Script de recherche de mots-clés dans les documents d'un répertoire

Auteur :
[**@kurama-a**](https://github.com/kurama-a)

## Description

Ce script permet d'extraire du texte et des images de divers types de fichiers (Word, Excel, PowerPoint, PDF, TXT...). Il recherche ensuite des mots-clés dans le texte extrait. Les résultats sont enregistrés dans un fichier csv.

### Types de fichiers pris en charge :
- `.docx`
- `.doc` 
- `.xlsx` 
- `.pptx` 
- `.pdf`
- `.txt` 
- `.csv`
- `.odt`
- `.ods`
- `.odp`
- `.xls`

## Prérequis

### 1. Installation des bibliothèques Python

- Python

Avant d'exécuter le script, assurez-vous d'installer les bibliothèques Python nécessaires via le fichier `requirements.txt` :

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Installation de Poppler

Poppler (pour la gestion des PDF) 

- **install_poppler.bat** : Installe Poppler et configure automatiquement le chemin dans votre système.

#### Instructions d'installation :

1. Localisez les fichiers batch :
   - `install_poppler.bat`

2. Exécutez **install_poppler.bat** en tant qu'administrateur :
   - Clic droit sur le fichier `.bat` → **Exécuter en tant qu’administrateur**.

Cela installera les outils nécessaires et les ajoutera au **PATH** de votre système.

# Utilisation

## 1. **Préparation des fichiers et répertoires**

- Le fichier **process.log** va contenir tous les fichiers testés avec leur emplacement, il va permettre de reprendre là où le programmme s'est arrêté en cas de problème.
- **Ajouter les mots-clés dans le fichier** nommé `keywords.txt` contenant une liste de mots-clés séparés par des virgules.

```
mot-clé1,mot-clé2,phrase test
```
Les mots-clés sont **insensibles à la casse** (`Mot-Clé` = `mot-clé`).


## 2. **Exécution du script** 

Lance le script en spécifiant le répertoire à analyser et le fichier de mots-clés :

```bash
python keyword_finder.py chemin_répertoire fichier_mots_clés.txt
```

### **Options disponibles :**
| Option | Description |
|--------|-------------|
| `--csv_output_file <nom_du_fichier>` | Spécifie un fichier CSV personnalisé pour stocker les résultats. |
| `--overwrite_logs` | Recherche des mots clés même dans les fichiers déjà scannés. |

Exemples :
```bash
python keyword_finder.py documents keywords.txt --csv_output_file resultats.csv

python keyword_finder.py documents keywords.txt --overwrite_logs
```

---

## 3. **Résultats de l’analyse**

- **Les résultats seront enregistrés dans un fichier CSV** (`keyword_search_results.csv` par défaut, ou celui que tu as spécifié avec `--csv_output_file`).
- **Chaque ligne du fichier CSV contient :**
  - **Répertoire** où le fichier a été trouvé.
  - **Nom du fichier**.
  - **URLs extraites** si des liens sont trouvés dans le fichier.






