# Script de recherche de mots-clés dans les documents d'un répertoire

## Description

Ce script permet d'extraire du texte et des images de divers types de fichiers (Word, Excel, PowerPoint, PDF, TXT, et images). Il recherche ensuite des mots-clés dans le texte extrait, ainsi que dans les images extraites via l'OCR. Les résultats sont enregistrés dans un fichier texte nommé `results.txt`.

### Types de fichiers pris en charge :
- `.docx` : Fichiers Word
- `.xlsx` : Fichiers Excel
- `.pptx` : Fichiers PowerPoint
- `.pdf` : Fichiers PDF
- `.txt` : Fichiers texte brut
- Images : `.jpg`, `.jpeg`, `.png`, `.bmp` (OCR via Tesseract)

## Prérequis

### 1. Installation des bibliothèques Python

Avant d'exécuter le script, assurez-vous d'installer les bibliothèques Python nécessaires via le fichier `requirements.txt` :

```bash
pip install -r requirements.txt
```

### 2. Installation de Poppler et Tesseract

Poppler (pour la gestion des PDF) et Tesseract (pour la reconnaissance optique des caractères) sont des outils externes nécessaires au bon fonctionnement du script. Pour simplifier l'installation, deux fichiers batch sont fournis :

- **install_poppler.bat** : Installe Poppler et configure automatiquement le chemin dans votre système.
- **install_tesseract.bat** : Installe Tesseract et le configure dans le PATH.

#### Instructions d'installation :

1. Téléchargez le projet et localisez les fichiers batch :
   - `install_poppler.bat`
   - `install_tesseract.bat`

2. Exécutez **chacun de ces fichiers** en tant qu'administrateur :
   - Clic droit sur le fichier `.bat` → **Exécuter en tant qu’administrateur**.

Cela installera les outils nécessaires et les ajoutera au **PATH** de votre système.

## Utilisation

### 1. Préparation des fichiers et répertoires :

- Placez les fichiers à analyser dans un répertoire (par exemple, `documents`).
- Créez un fichier texte `keyword.txt` dans lequel chaque ligne contient un mot-clé à rechercher dans les documents.

### 2. Exécution du script :

Pour exécuter le script, vous devez spécifier le répertoire contenant les fichiers à analyser. Vous pouvez le faire en passant le chemin du dossier comme argument lors de l'exécution du script.

Par exemple, pour analyser un dossier appelé `documents` :

```bash
python script.py documents
```

### 3. Résultats :

- Les résultats de la recherche de mots-clés seront sauvegardés dans un fichier texte nommé `results.txt` dans le même répertoire que le script.
- Les images extraites des documents seront enregistrées dans un répertoire appelé `images`.

### 4. Fichier de mots-clés :

Le fichier `keyword.txt` doit contenir un mot-clé par ligne, que le script utilisera pour rechercher ces mots dans les documents.

Exemple de contenu de `keyword.txt` :

```
mot-clé1
mot-clé2
mot-clé3
```
