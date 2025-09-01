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

Here is the English translation:

---

## 1. **Preparation of Files and Directories**

- The **process.log** file will contain all tested files with their locations. It allows the program to resume where it left off in case of an issue.
- **Add the keywords to a file** named `keywords.txt` containing a list of keywords separated by commas.

```
keyword1,keyword2,test phrase
```
Keywords are **case-insensitive** (`Keyword` = `keyword`).

## 2. **Running the Script**

Run the script by specifying the directory to analyze and the keywords file:

```bash
python keyword_finder.py directory_path keywords_file.txt
```

### **Available Options:**
| Option | Description |
|--------|-------------|
| `--csv_output_file <file_name>` | Specifies a custom CSV file to store the results. |
| `--overwrite_logs` | Searches for keywords even in already scanned files. |

Examples:
```bash
python keyword_finder.py documents keywords.txt --csv_output_file results.csv

python keyword_finder.py documents keywords.txt --overwrite_logs
```

## 3. **Résultats de l’analyse**

- **Les résultats seront enregistrés dans un fichier CSV** (`keyword_search_results.csv` par défaut, ou celui que tu as spécifié avec `--csv_output_file`).
- **Chaque ligne du fichier CSV contient :**
  - **Répertoire** où le fichier a été trouvé.
  - **Nom du fichier**.
  - **URLs extraites** si des liens sont trouvés dans le fichier.






