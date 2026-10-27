import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont
import os
import csv

# --- CONFIGURATION ---
INPUT_FILE = "Renouvellement.pdf"
EMPLOYEES_FILE = "liste_employes.csv" 
OUTPUT_DIR = "Dossiers_Generes"

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

FONT_SIZE = 22
H_PAGE = 842

# --- CORRECTIONS GLOBALES ---
CORRECTION_X = -5
CORRECTION_Y = 5

def cy(pdf_y):
    return H_PAGE - pdf_y - 8

DATA_FIXE = [
        (0, 521, 646, "1h00"),
    (0, 377, 542, "EQUIPIER"),
    (0, 153, 549, "nom"),
    (0, 110, 292, "X"),
    (0, 262, 230, "X"),
    (0, 253, 180, "X"),
    (0, 320, 167, "X"),
    (0, 391, 144, "X"),
    (0, 394, 100, "X"),
    (1, 112, 445, "X"),
    (1, 117, 328, "X"),
    (1, 110, 110, "X"),
    (2, 256, 773, "X"),
    (2, 325, 723, "X"),
    (2, 110, 643, "X"),
    (2, 110, 630, "X"),
    (2, 107, 548, "X"),
    (2, 110, 454, "X"),
    (2, 397, 362, "X"),
    (2, 250, 313, "X"),
    (2, 250, 313, "X"),
    (2, 109, 241, "X"),
    (2, 110, 176, "X"),
    (2, 111, 151, "X"),
    (2, 110, 84, "X"),
    (3, 108, 774, "X"),
    (3, 110, 626, "X"),
    (3, 108, 643, "X"),
    (4, 82, 503, "X"),
    (4, 230, 439, "X"),
    (4, 73, 325, "X"),
    (4, 84, 249, "X"),
    (4, 77, 199, "X"),
    (4, 79, 84, "X"),
    (5, 83, 721, "X"),
    (5, 85, 550, "X"),
    (5, 92, 414, "X"),
    (5, 87, 330, "X"),
    (5, 85, 245, "X"),
    (5, 85, 79, "X"),
    (6, 278, 542, "X"),
    (6, 346, 511, "X"),
    (6, 397, 487, "X"),
    (6, 82, 760, "X"),
    (6, 90, 640, "X"),
    (6, 82, 610, "X"),
    (6, 87, 483, "X"),
    (6, 85, 376, "X"),
    (6, 87, 284, "X"),
    (6, 85, 172, "X"),
    (7, 82, 709, "X"),
    (7, 75, 580, "X"),
    (7, 77, 459, "X"),
    (6, 84, 756, "X"),
    (6, 290, 536, "X"),
    (6, 342, 512, "X"),
    (6, 395, 483, "X"),
    (6, 84, 636, "X"),
    (6, 86, 384, "X"),
    (6, 81, 283, "X"),
    (6, 80, 168, "X"),
    (7, 79, 716, "X"),
    (7, 71, 575, "X"),
    (7, 77, 506, "X"),
    (8, 394, 761, "nom"),
    (8, 105, 800, "date"),
    (9, 20, 642, "X"),
    (9, 19, 593, "X"),
    (9, 22, 499, "X"),
    (9, 20, 412, "X"),
    (9, 17, 320, "X"),
    (9, 20, 237, "X"),
    (10, 26, 772, "X"),
    (10, 18, 689, "X"),
    (10, 20, 545, "X"),
    (10, 22, 412, "X"),
    (10, 19, 309, "X"),
    (11, 448, 726, "nom"),
    (11, 115, 756, "date"),

]

def generate_one_pdf(nom, date, poste, output_path):
    doc = fitz.open(INPUT_FILE)
    output_images = []
    
    try:
        font = ImageFont.truetype("arial.ttf", FONT_SIZE)
    except:
        font = ImageFont.load_default()

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        draw = ImageDraw.Draw(img)

        
        items = [i for i in DATA_FIXE if i[0] == page_num]
        
        for item in items:
            _, x_pdf, y_pdf, text_raw = item
            
            
            text_to_draw = text_raw
            if text_raw == "nom": text_to_draw = nom
            elif text_raw == "date": text_to_draw = date
            elif text_raw == "poste": text_to_draw = poste
            
            # Positionnement
            x_img = (x_pdf * 2) + CORRECTION_X
            y_img = (cy(y_pdf) * 2) - CORRECTION_Y 
            
            draw.text((x_img, y_img), text_to_draw, font=font, fill=(0, 0, 0))
        
        output_images.append(img)

    output_images[0].save(output_path, save_all=True, append_images=output_images[1:], resolution=150.0, quality=90)

# --- LANCEMENT ---
if __name__ == "__main__":
    if not os.path.exists(EMPLOYEES_FILE):
        print(f"Erreur : '{EMPLOYEES_FILE}' introuvable.")
    else:
        with open(EMPLOYEES_FILE, newline='', encoding='latin-1') as f:
            reader = csv.reader(f, delimiter=',') 
            try:
                next(reader) 
            except StopIteration:
                pass

            for row in reader:
                if len(row) >= 2:
                    nom = row[0].strip().upper()
                    date = row[1].strip()
                    poste = row[2].strip().upper() if len(row) > 2 else "EQUIPIER"
                    
                    outfile = os.path.join(OUTPUT_DIR, f"Fiche_{nom}.pdf")
                    print(f"Génération de la fiche pour {nom}...")
                    
                    generate_one_pdf(nom, date, poste, outfile)
        
        print(f"\nTerminé ! Vos fichiers sont dans '{OUTPUT_DIR}'.")