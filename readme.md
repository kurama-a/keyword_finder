# Script de recherche de mots-clés dans les documents d'un répertoire

## Description

Ce script permet d'extraire du texte et des images de divers types de fichiers (Word, Excel, PowerPoint, PDF, TXT, et images). Il recherche ensuite des mots-clés dans le texte extrait, ainsi que dans les images extraites via l'OCR. Les résultats sont enregistrés dans un fichier texte nommé `results.txt`.

### Types de fichiers pris en charge :
- `.docx` : Fichiers Word
- `.doc` : Fichiers Word
- `.xlsx` : Fichiers Excel
- `.pptx` : Fichiers PowerPoint
- `.pdf` : Fichiers PDF
- `.txt` : Fichiers texte brut

## Prérequis

### 1. Installation des bibliothèques Python

Avant d'exécuter le script, assurez-vous d'installer les bibliothèques Python nécessaires via le fichier `requirements.txt` :

```bash
pip install -r requirements.txt
```

### 2. Installation de Poppler

Poppler (pour la gestion des PDF) 

- **install_poppler.bat** : Installe Poppler et configure automatiquement le chemin dans votre système.

#### Instructions d'installation :

1. Téléchargez le projet et localisez les fichiers batch :
   - `install_poppler.bat`

2. Exécutez **chacun de ces fichiers** en tant qu'administrateur :
   - Clic droit sur le fichier `.bat` → **Exécuter en tant qu’administrateur**.

Cela installera les outils nécessaires et les ajoutera au **PATH** de votre système.

## Utilisation

### 1. Préparation des fichiers et répertoires :

- Placez les fichiers à analyser dans un répertoire (par exemple, `documents`).
- Créez un fichier texte `keyword.txt` dans lequel tous les mot-clé à rechercher sont séparés par une virgule.

### 2. Exécution du script :

Pour exécuter le script, vous devez spécifier le répertoire contenant les fichiers à analyser. Vous pouvez le faire en passant le chemin du dossier comme argument lors de l'exécution du script.

Par exemple, pour analyser un dossier appelé `documents` si vous êtes dans le répertoire :

```bash
python keyword_finder.py documents keyword.txt
```

### 3. Résultats :

- Les résultats de la recherche de mots-clés seront sauvegardés dans un fichier texte nommé `keyword_search_results.csv` dans le même répertoire que le script.

### 4. Fichier de mots-clés :

Le fichier `keyword.txt` doit contenir tous les mots clés séparés par des virgules.

Exemple de contenu de `keyword.txt` :

```
mot-clé1,mot-clé2,phrase test
```
