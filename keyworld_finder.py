import os
import pandas as pd
from docx import Document
from pptx import Presentation
import PyPDF2
from PIL import Image
import pytesseract
from pdf2image import convert_from_path

# Extraction du texte d'un fichier Word
def extract_text_from_docx(file_path, output_folder):
    doc = Document(file_path)
    text = ' '.join([para.text for para in doc.paragraphs])
    
    # Extraction des images
    for i, rel in enumerate(doc.part.rels.values()):
        if "image" in rel.target_ref:
            img_data = rel.target_part.blob
            img_filename = os.path.join(output_folder, f"{os.path.basename(file_path)}_image_{i}.png")
            with open(img_filename, "wb") as img_file:
                img_file.write(img_data)
            print(f"Image extraite: {img_filename}")
    
    return text

# Extraction du texte d'un fichier Excel
def extract_text_from_excel(file_path):
    data = pd.read_excel(file_path)
    return ' '.join(data.astype(str).values.flatten())

# Extraction du texte d'un fichier PowerPoint
def extract_text_from_pptx(file_path, output_folder):
    prs = Presentation(file_path)
    text = ''
    for i, slide in enumerate(prs.slides):
        for j, shape in enumerate(slide.shapes):
            if hasattr(shape, "text"):
                text += shape.text + ' '
            # Extraction des images
            if hasattr(shape, "image"):
                img = shape.image
                img_filename = os.path.join(output_folder, f"{os.path.basename(file_path)}_image_{i}_{j}.png")
                with open(img_filename, "wb") as img_file:
                    img_file.write(img.blob)
                print(f"Image extraite: {img_filename}")
    return text

# Extraction du texte d'un fichier PDF et des images
def extract_text_from_pdf(file_path, output_folder):
    text = ''
    with open(file_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text()

    # Convertir les pages PDF en images et traiter avec OCR
    images = convert_from_path(file_path)
    for i, image in enumerate(images):
        img_filename = os.path.join(output_folder, f"{os.path.basename(file_path)}_image_{i}.png")
        image.save(img_filename, 'PNG')
        print(f"Image extraite: {img_filename}")
    
    return text

# Extraction du texte d'un fichier TXT
def extract_text_from_txt(file_path):
    with open(file_path, 'r') as file:
        return file.read()

# Extraction du texte d'une image (OCR)
def extract_text_from_image(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img)
    return text

# Fonction pour détecter le type de fichier et extraire le texte correspondant
def extract_text(file_path, output_folder):
    ext = os.path.splitext(file_path)[-1].lower()
    if ext == '.docx':
        return extract_text_from_docx(file_path, output_folder)
    elif ext == '.xlsx':
        return extract_text_from_excel(file_path)
    elif ext == '.pptx':
        return extract_text_from_pptx(file_path, output_folder)
    elif ext == '.pdf':
        return extract_text_from_pdf(file_path, output_folder)
    elif ext == '.txt':
        return extract_text_from_txt(file_path)
    elif ext in ['.jpg', '.jpeg', '.png', '.bmp']:
        return extract_text_from_image(file_path)
    else:
        raise ValueError(f"Unsupported file format: {ext}")

# Fonction pour rechercher les mots-clés dans le texte extrait
def find_keywords_in_text(text, keywords):
    found_keywords = [keyword for keyword in keywords if keyword.lower() in text.lower()]
    return found_keywords

# Fonction pour rechercher les mots-clés dans plusieurs fichiers, y compris les sous-répertoires, et écrire les résultats dans un fichier
def search_keywords_in_files(directory_path, keywords, output_folder, output_file):
    with open(output_file, 'w') as result_file:
        # Parcourir récursivement les fichiers dans le répertoire et ses sous-répertoires
        for root, dirs, files in os.walk(directory_path):
            for file_name in files:
                file_path = os.path.join(root, file_name)
                
                try:
                    # Extraire le texte du fichier
                    text = extract_text(file_path, output_folder)
                    
                    # Rechercher les mots-clés dans le texte
                    found_keywords = find_keywords_in_text(text, keywords)
                    
                    # Rechercher les mots-clés dans les images extraites
                    images_found_keywords = False
                    for img_file in os.listdir(output_folder):
                        if img_file.startswith(os.path.basename(file_path)):
                            img_path = os.path.join(output_folder, img_file)
                            if os.path.isfile(img_path):
                                ocr_text = extract_text_from_image(img_path)
                                found_keywords_in_image = find_keywords_in_text(ocr_text, keywords)
                                if found_keywords_in_image:
                                    images_found_keywords = True

                    # Si des mots-clés sont trouvés dans le fichier ou ses images, ajouter au résultat
                    if found_keywords or images_found_keywords:
                        result_file.write(f"File: {file_path}, ")  # Écrire le chemin complet du fichier
                        if found_keywords:
                            result_file.write(f"Keyword found in the text: \"{', '.join(found_keywords)}\"")
                            if images_found_keywords == True:
                                result_file.write(f" | ")
                        if images_found_keywords:
                            result_file.write(f"Keyword found in the image \"{', '.join(found_keywords_in_image)}\" ")
                        result_file.write("\n")
                except Exception as e:
                    print(f"Error processing {file_name}: {e}")
                    
def keyword_file(keyword_file_path):
    with open(keyword_file_path, "r", encoding="utf8") as keyword_file:
        keyword_list = []   
        for line in keyword_file:
            keyword_list.append(line.strip()) 
    return keyword_list
        
# Exemple d'utilisation
directory_path = 'documents'  
output_folder = 'images'  # Répertoire pour sauvegarder les images extraites
output_file = 'results.txt'  # Fichier pour enregistrer les résultats
keyword_file_path = "keyword.txt"
os.makedirs(output_folder, exist_ok=True)  # Crée le répertoire d'images s'il n'existe pas
search_keywords_in_files(directory_path, keyword_file(keyword_file_path), output_folder, output_file)
