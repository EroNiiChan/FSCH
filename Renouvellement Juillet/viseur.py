import fitz  # PyMuPDF
import matplotlib.pyplot as plt
import numpy as np

# --- CONFIGURATION ---
INPUT_FILE = "Renouvellement.pdf"
OUTPUT_TXT = "coordonnees_turbo_v2.txt"
H_PAGE = 842

# Variables d'état
current_page_index = 0
current_type = "X" # Par défaut
doc = None
ax = None
fig = None

def get_variable_name():
    if current_type == "X": return '"X"'
    if current_type == "NOM": return 'nom'
    if current_type == "DATE": return 'date'
    if current_type == "FONCTION": return '"EQUIPIER"' # Valeur par défaut modifiable
    if current_type == "DUREE": return '"1h00"'       # Valeur par défaut modifiable
    return '"X"'

def refresh_display():
    """Recharge l'image de la page actuelle"""
    global ax, fig
    ax.clear()
    
    page = doc.load_page(current_page_index)
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
    img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    
    ax.imshow(img_array)
    
    # Titre dynamique
    titre = f"PAGE {current_page_index + 1} | Mode actuel : {current_type}"
    titre += "\n[<-- / -->] Changer page"
    titre += "\n[1] Croix  [2] Nom  [3] Date  [4] Fonction  [5] Durée"
    ax.set_title(titre, fontsize=12, color='blue', fontweight='bold')
    
    fig.canvas.draw()

def on_key(event):
    """Gestion du clavier"""
    global current_page_index, current_type
    
    if event.key == 'right':
        if current_page_index < len(doc) - 1:
            current_page_index += 1
            refresh_display()
            print(f"\n---> Passage à la page {current_page_index + 1}")
    
    elif event.key == 'left':
        if current_page_index > 0:
            current_page_index -= 1
            refresh_display()
            print(f"\n<--- Retour à la page {current_page_index + 1}")
            
    elif event.key == '1':
        current_type = "X"
        refresh_display()
        print("Mode : CROIX (X)")
        
    elif event.key == '2':
        current_type = "NOM"
        refresh_display()
        print("Mode : NOM")
        
    elif event.key == '3':
        current_type = "DATE"
        refresh_display()
        print("Mode : DATE")

    elif event.key == '4':
        current_type = "FONCTION"
        refresh_display()
        print("Mode : FONCTION")

    elif event.key == '5':
        current_type = "DUREE"
        refresh_display()
        print("Mode : DUREE")

def on_click(event):
    """Gestion du clic souris"""
    if event.xdata is None or event.ydata is None:
        return

    # Calcul des coordonnées
    final_x = int(event.xdata / 2)
    final_y = int(H_PAGE - (event.ydata / 2) - 8)

    
    valeur = get_variable_name()
    ligne_code = f"    ({current_page_index}, {final_x}, {final_y}, {valeur}),"
    
    
    with open(OUTPUT_TXT, "a", encoding="utf-8") as f:
        f.write(ligne_code + "\n")
    
    
    couleur = 'ro' # Rouge par défaut
    if current_type == "NOM": couleur = 'bo' # Bleu
    if current_type == "FONCTION": couleur = 'go' # Vert
    
    ax.plot(event.xdata, event.ydata, couleur, markersize=5)
    fig.canvas.draw()
    
    print(f"Enregistré ({current_type} - P{current_page_index+1}) : {ligne_code}")

def main():
    global doc, fig, ax
    
    try:
        doc = fitz.open(INPUT_FILE)
    except:
        print(f"Impossible d'ouvrir {INPUT_FILE}")
        return

    print("--- VISEUR TURBO V2 ---")
    print("Commandes Clavier :")
    print("  [1] Croix (X)")
    print("  [2] Nom")
    print("  [3] Date")
    print("  [4] Fonction (Ajouté !)")
    print("  [5] Durée (Ajouté !)")
    print("  [Flèches] Changer de page")
    print(f"Sauvegarde dans : '{OUTPUT_TXT}'")

    fig, ax = plt.subplots(figsize=(12, 16))
    refresh_display()
    
    fig.canvas.mpl_connect('button_press_event', on_click)
    fig.canvas.mpl_connect('key_press_event', on_key)
    
    plt.show()

if __name__ == "__main__":
    main()