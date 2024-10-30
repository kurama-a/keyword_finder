import os
import re
import pandas as pd
import argparse
from docx import Document
from pptx import Presentation
import PyPDF2
import win32com.client as win32
from concurrent.futures import ProcessPoolExecutor

PROGRESS_LOG_FILE = 'progress.log'
ERROR_LOG_FILE = 'error.log'

# Extraction du texte d'un fichier Word
def extract_text_from_docx(file_path):
    doc = Document(file_path)
    text = ' '.join([para.text for para in doc.paragraphs])
    return text

# Fonction pour extraire le texte d'un fichier Word (.doc)
def extract_text_from_doc(file_path):
    try:
        word = win32.gencache.EnsureDispatch("Word.Application")
        doc = word.Documents.Open(file_path)
        text = doc.Content.Text
        doc.Close()
        word.Quit()
        return text
    except ImportError:
        raise ImportError("pywin32 doit être installé pour extraire le texte d'un fichier .doc")

# Extraction du texte d'un fichier Excel
def extract_text_from_excel(file_path):
    data = pd.read_excel(file_path)
    return ' '.join(data.astype(str).values.flatten())

# Extraction du texte d'un fichier PowerPoint
def extract_text_from_pptx(file_path):
    prs = Presentation(file_path)
    text = ''
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + ' '
    return text

# Extraction du texte d'un fichier PDF
def extract_text_from_pdf(file_path):
    text = ''
    with open(file_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text()
    return text

# Extraction du texte d'un fichier TXT
def extract_text_from_txt(file_path):
    with open(file_path, 'r') as file:
        return file.read()

# Fonction pour détecter les URLs dans le texte extrait
def find_urls(text):
    url_pattern = r'(https?://[^\s]+)'
    urls = re.findall(url_pattern, text)
    return urls

# Fonction pour détecter le type de fichier et extraire le texte correspondant
def extract_text(file_path):
    ext = os.path.splitext(file_path)[-1].lower()
    if ext == '.docx':
        return extract_text_from_docx(file_path)
    elif ext == '.doc':
        return extract_text_from_doc(file_path)
    elif ext == '.xlsx':
        return extract_text_from_excel(file_path)
    elif ext == '.pptx':
        return extract_text_from_pptx(file_path)
    elif ext == '.pdf':
        return extract_text_from_pdf(file_path)
    elif ext == '.txt':
        return extract_text_from_txt(file_path)
    else:
        raise ValueError(f"Unsupported file format: {ext}")

# Fonction pour rechercher les mots-clés et les URLs dans le texte extrait
def find_keywords_and_urls_in_text(text, keywords):
    found_keywords = [keyword for keyword in keywords if keyword.lower() in text.lower()]
    urls = find_urls(text)
    return found_keywords, urls

# Fonction pour traiter un fichier et rechercher les mots-clés
def process_file(file_info, keywords):
    file_path, root_dir = file_info
    try:
        text = extract_text(file_path)
        found_keywords, urls = find_keywords_and_urls_in_text(text, keywords)
        
        if found_keywords:
            return {
                'Directory': os.path.abspath(root_dir), 
                'Filename': os.path.basename(file_path),
                'URLs': ';'.join(set(urls))
            }
    except Exception as e:
        # Enregistrer les erreurs dans error.log
        with open(ERROR_LOG_FILE, 'a') as error_file:
            error_file.write(f"Error processing {file_path}: {e}\n")
    return None


# Fonction pour lire l'état de progression
def load_progress():
    if os.path.exists(PROGRESS_LOG_FILE):
        with open(PROGRESS_LOG_FILE, 'r') as f:
            return set(line.strip() for line in f)
    return set()

# Fonction pour enregistrer l'état de progression en mode 'append'
def save_progress(processed_files):
    with open(PROGRESS_LOG_FILE, 'a') as f:
        for file in processed_files:
            f.write(f"{file}\n")

# Fonction principale pour rechercher les mots-clés en parallèle
def search_keywords_in_files(directory_path, keywords, csv_output_file, batch_size=100):
    results_data = []
    processed_files = load_progress()
    file_paths = [(os.path.join(root_dir, file_name), root_dir) 
                  for root_dir, _, files in os.walk(directory_path) 
                  for file_name in files if os.path.join(root_dir, file_name) not in processed_files]

    with ProcessPoolExecutor() as executor:
        futures = []
        batch = []
        
        # Traiter les fichiers en lot pour sauvegarder périodiquement
        for file_info in file_paths:
            batch.append(file_info)
            if len(batch) >= batch_size:
                futures.append(executor.submit(process_batch, batch, keywords))
                batch = []

        # Traiter les fichiers restants
        if batch:
            futures.append(executor.submit(process_batch, batch, keywords))
        
        # Collecte des résultats
        for future in futures:
            batch_results, batch_files = future.result()
            results_data.extend(batch_results)
            processed_files.update(batch_files)
            save_progress(batch_files)

    # Sauvegarder les résultats finaux dans le fichier CSV en mode 'append'
    if results_data:
        df = pd.DataFrame(results_data)
        df.to_csv(csv_output_file, index=False, mode='a', header=not os.path.exists(csv_output_file))

    print(f"Scan completed. Results saved to {csv_output_file}")

# Fonction pour traiter un lot de fichiers
def process_batch(file_batch, keywords):
    batch_results = []
    batch_files = []

    for file_info in file_batch:
        result = process_file(file_info, keywords)
        if result:
            batch_results.append(result)
            batch_files.append(file_info[0])

    return batch_results, batch_files

# Fonction pour charger les mots-clés à partir d'un fichier
def load_keywords(keyword_file):
    with open(keyword_file, 'r') as file:
        return file.read().strip().split(',')

# Fonction principale
def main():
    parser = argparse.ArgumentParser(description="Recherche de mots-clés dans des fichiers.")
    parser.add_argument('directory', type=str, help="Chemin du répertoire à scanner.")
    parser.add_argument('keyword_file', type=str, help="Fichier contenant les mots-clés séparés par des virgules.")
    parser.add_argument('--csv_output_file', type=str, default='keyword_search_results.csv', help="Fichier CSV pour enregistrer les résultats.")
    args = parser.parse_args()

    # Charger les mots-clés depuis le fichier
    keywords = load_keywords(args.keyword_file)

    # Recherche des mots-clés en parallèle et enregistrement des résultats
    search_keywords_in_files(args.directory, keywords, args.csv_output_file)

if __name__ == "__main__":
    main()
