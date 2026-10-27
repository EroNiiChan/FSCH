#!/usr/bin/env python3
"""
Remplissage automatique des fiches d'évaluation et de renouvellement (Intégration / Hygiène / Sécurité).
- Entrée : Fichier Excel (.xlsx / .xls) ou CSV détecté automatiquement.
- Prise en charge des cellules fusionnées et nettoyage des sauts de ligne internes (\\n).
- Ajustement automatique de la taille de police (auto-fit) pour les noms longs.
- Correction exacte et calée au pixel près pour chaque exercice :
  * Exercice 5 : Numérotation dans l'ordre des 7 étapes (4, 1, 7, 5, 2, 6, 3).
  * FCH, FCS, FCAPH, FCGP : Croix placées exactement au centre des cases des bonnes réponses.
  * Gestion des pages US Letter (hauteur 792 pt au milieu du formulaire) pour éviter tout décalage sous les cases.
  * Validation OUI sur toutes les fiches d'évaluation.
- Préservation des annotations manuscrites et tirage aléatoire parmi les 3 modèles.
"""

import os
import sys
import glob
import re
import random
import argparse
import time
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import fitz  # PyMuPDF

# Coordonnées des champs texte personnalisés (page_idx, x, y_baseline, clé, largeur_max_pt)
TEXT_FIELDS = [
    # Page 1 (FCH)
    (0, 75, 298, "nom", 235),
    (0, 318, 298, "poste", 150),
    (0, 368, 224, "date", 120),
    # Page 9 (FCS)
    (8, 338, 74, "nom", 210),
    (8, 55, 36, "date", 120),
    # Page 12 (FCAPH - US Letter h=792)
    (11, 79, 407, "nom", 295),
    (11, 383, 407, "poste", 150),
    (11, 393, 284, "date", 120),
    # Page 16 (FCAPH suite)
    (15, 370, 84, "nom", 210),
    (15, 85, 45, "date", 120),
    # Page 18 (FCGP)
    (17, 374, 503, "nom", 210),
    (17, 63, 460, "date", 120),
]

# --- EXERCICE 5 (Page 2, index 1) : Numérotation des 7 étapes de nettoyage et d'aseptisation ---
# 1. Aseptisation (étape 4)
# 2. Prélavage à l'eau chaude (étape 1)
# 3. Entreposage sur étagère (étape 7)
# 4. Rinçage à l'eau froide (étape 5)
# 5. Nettoyage à l'eau chaude (étape 2)
# 6. Séchage à l'air libre (étape 6)
# 7. Rinçage à l'eau chaude (étape 3)
EXO_5_NUMBERS = [
    (1, 111.0, 68.0, "4"),
    (1, 111.0, 101.0, "1"),
    (1, 111.0, 134.5, "7"),
    (1, 111.0, 168.0, "5"),
    (1, 111.0, 201.0, "2"),
    (1, 111.0, 234.5, "6"),
    (1, 111.0, 267.5, "3"),
]

# --- COORDONNÉES EXACTES ET VÉRIFIÉES DES BONNES RÉPONSES (Croix "X") ---
CORRECTION_CROSSES = [
    # ---------------- PAGE 1 (FCH - Hygiène) ----------------
    (0, 74.0, 227.0),   # Formation de suivi
    (0, 109.0, 541.5),  # Q1 : Installer un panneau de signalisation
    (0, 251.0, 607.0),  # Q2 : Tiède
    (0, 251.0, 660.0),  # Q3 : Bac en plonge -> 4h00
    (0, 251.0, 673.5),  # Q3 : Seaux à chiffons -> 4h00
    (0, 463.0, 687.0),  # Q3 : Vaporisateurs -> 24h00
    (0, 392.5, 740.0),  # Q4 : Toutes les 4h00

    # ---------------- PAGE 2 (FCH Suite) ----------------
    # (L'Exercice 5 est géré par EXO_5_NUMBERS ci-dessus)
    (1, 109.0, 391.0),  # Q6 : Dès que nécessaire
    (1, 109.0, 404.0),  # Q6 : Au minimum 1 fois par heure
    (1, 110.0, 552.0),  # Q7 : Savon bactéricide jusqu'aux avant-bras 30s
    (1, 109.0, 751.0),  # Q8 : Nettoyer, désinfecter puis protection étanche

    # ---------------- PAGE 3 (FCH Suite) ----------------
    (2, 251.0, 63.0),   # Q9 : Mon alliance
    (2, 428.0, 116.0),  # Q10 : Le local poubelle (intrus)
    (2, 109.0, 196.0),  # Q11 : Faux ongles et/ou vernis
    (2, 109.0, 209.0),  # Q11 : Tenue de travail hors zone
    (2, 109.0, 289.0),  # Q12 : Un gant gril dédié
    (2, 109.0, 302.5),  # Q12 : Une pince
    (2, 109.0, 382.5),  # Q13 : Vaporiser sur un chiffon propre
    (2, 392.5, 475.5),  # Q14 : Entre 1 et 4°C
    (2, 251.0, 528.5),  # Q15 : Entre -18 et -23°C
    (2, 109.0, 595.0),  # Q16 : Jeter les tomates et les noter en perte
    (2, 109.0, 661.5),  # Q17 : PEPS
    (2, 109.0, 688.0),  # Q17 : Les vérifier chaque jour
    (2, 109.0, 753.5),  # Q18 : Pour éliminer toute trace d'humidité

    # ---------------- PAGE 4 (FCH Suite) ----------------
    (3, 109.0, 63.0),   # Q19 : Pour éliminer toute trace de nourriture
    (3, 109.0, 196.0),  # Q20 : En nettoyant et aseptisant à intervalle régulier
    (3, 109.0, 209.0),  # Q20 : En respectant les procédures de cuisson

    # ---------------- PAGE 5 (FCS - Sécurité) ----------------
    (4, 164.0, 201.0),  # Formation de renouvellement
    (4, 75.0, 329.0),   # Q1 : Décrocher / Dégoupiller / Percuter / Attaquer
    (4, 218.0, 398.0),  # Q2 : 18 (pompiers)
    (4, 285.0, 398.0),  # Q2 : 112 (urgences)
    (4, 75.0, 508.0),   # Q3 : Celui qui se trouve à proximité de la zone en feu
    (4, 75.0, 633.0),   # Q4 : Ne plus découvrir la friteuse
    (4, 76.0, 755.0),   # Q5 : Jamais (issue de secours bloquée)

    # ---------------- PAGE 6 (FCS Suite) ----------------
    (5, 77.0, 119.0),   # Q6 : Méthode échelonnée (l'ensemble des équipiers...)
    (5, 78.0, 281.0),   # Q7 : Le personnel planifié
    (5, 77.0, 417.0),   # Q8 : Te mettre en retrait et ne plus intervenir
    (5, 78.0, 504.0),   # Q9 : Un mètre au minimum
    (5, 77.0, 589.0),   # Q10 : Ne pas jouer au héros...
    (5, 79.0, 751.0),   # Q11 : Visière / tablier / gants

    # ---------------- PAGE 7 (FCS Suite) ----------------
    (6, 78.0, 79.0),    # Q12 : Signaler au responsable de quart
    (6, 78.0, 196.0),   # Q13 : Couper le courant, dégager la victime...
    (6, 287.0, 299.0),  # Q14 : Equipière mineure (10 kg)
    (6, 338.0, 324.0),  # Q14 : Equipier mineur (20 kg)
    (6, 338.0, 347.0),  # Q14 : Equipière majeure (20 kg)
    (6, 79.0, 431.0),   # Q15 : Sur l'étiquette et dans le classeur Sécurité
    (6, 79.0, 453.0),   # Q15 : Sur les consignes de Sécurité et fiches de données
    (6, 79.0, 552.0),   # Q16 : Compacteur, broyeur
    (6, 79.0, 666.0),   # Q17 : Vêtement fluorescent à bandes réfléchissantes

    # ---------------- PAGE 8 (FCS Suite) ----------------
    (7, 72.0, 124.0),   # Q18 : Ne pas pénétrer dans la zone de danger...
    (7, 71.0, 255.0),   # Q19 : S'échapper ou se cacher, et alerter
    (7, 72.0, 326.0),   # Q20 : Méthode SAFE

    # ---------------- PAGE 9 (Validation FCS) ----------------
    (8, 224.0, 170.0),  # FCS validée à 100% : OUI

    # ---------------- PAGE 12 (FCAPH - Safe & Inclusive, US Letter h=792) ----------------
    (11, 59.5, 287.5),  # Formation de suivi
    (11, 63.5, 548.5),  # Q1 : L'origine
    (11, 63.5, 562.5),  # Q1 : Le sexe
    (11, 63.5, 576.5),  # Q1 : Les croyances
    (11, 307.0, 548.5), # Q1 : La situation de handicap
    (11, 307.0, 576.5), # Q1 : L'âge
    (11, 63.5, 614.5),  # Q2 : Brimades régulières injustifiées
    (11, 63.5, 628.5),  # Q2 : Actes d'humiliation
    (11, 63.5, 642.5),  # Q2 : Mise à l'écart d'un salarié
    (11, 63.5, 680.5),  # Q3 : Allusions sexuelles répétées...
    (11, 63.5, 694.5),  # Q3 : Propos obscènes à l'oral ou l'écrit
    (11, 63.5, 708.5),  # Q3 : Sollicitations subies et répétées...
    (11, 63.5, 722.5),  # Q3 : Gestes déplacés... contacts physiques

    # ---------------- PAGE 13 (FCAPH Suite, US Letter h=792) ----------------
    (12, 63.5, 81.0),   # Q4 : Réduire un homme ou femme...
    (12, 63.5, 95.0),   # Q4 : Poster systématiquement femmes service...
    (12, 63.5, 109.0),  # Q4 : Blagues sexistes
    (12, 63.5, 147.0),  # Q5 : Les managers
    (12, 63.5, 161.0),  # Q5 : Les équipiers
    (12, 307.0, 147.0), # Q5 : Les clients
    (12, 307.0, 161.0), # Q5 : Les collègues
    (12, 63.5, 212.5),  # Q6 : Pour protéger la santé physique
    (12, 63.5, 226.5),  # Q6 : Pour protéger la santé mentale
    (12, 63.5, 240.5),  # Q6 : Climat serein
    (12, 63.5, 268.0),  # Q7 : Défendre les valeurs de notre entreprise
    (12, 63.5, 282.0),  # Q7 : Interdit par la loi
    (12, 63.5, 334.0),  # Q8 : En parler au responsable...
    (12, 63.5, 413.5),  # Q9 : En étant soi-même exemplaire
    (12, 63.5, 427.5),  # Q9 : Ecrits, paroles et actes respectueux
    (12, 63.5, 507.0),  # Q10 : Sanctions disciplinaires
    (12, 63.5, 521.0),  # Q10 : Sanctions judiciaires

    # ---------------- PAGE 14 (Accueil Handicap) ----------------
    (13, 76.5, 202.5),  # Q1 : L'accueillir comme n'importe quel client
    (13, 76.5, 249.5),  # Q1 : Traité de la même manière que tout autre client
    (13, 76.5, 391.0),  # Q2 : Mode handicapé sur la borne
    (13, 76.5, 529.0),  # Q3 : Proposer de s'installer à table
    (13, 76.5, 706.0),  # Q4 : S'installer en salle, service à table

    # ---------------- PAGE 15 (Accueil Handicap Suite) ----------------
    (14, 78.5, 108.0),  # Q5 : Chien guide d'aveugle -> guider jusqu'à table
    (14, 78.5, 219.0),  # Q6 : Proposer de rester à ses côtés à la borne
    (14, 78.5, 322.0),  # Q7 : Se présenter dame puis enfant avec bienveillance
    (14, 78.5, 502.5),  # Q8 : Prendre le temps d'écouter sans impatience
    (14, 78.5, 644.5),  # Q9 : Ecouter avec attention et patience
    (14, 78.5, 697.0),  # Q9 : Feuille et stylo pour communiquer
    (14, 78.5, 760.0),  # Q10 : Face à lui pour lecture sur les lèvres

    # ---------------- PAGE 16 (Validation FCAPH) ----------------
    (15, 224.0, 170.0), # FCAPH validée : OUI

    # ---------------- PAGE 17 (Gestes & Postures) ----------------
    (16, 92.5, 269.0),  # Q1 : Encadrer la charge avec ses pieds
    (16, 92.5, 399.0),  # Q2 : Plier les jambes et garder le dos droit
    (16, 92.5, 529.0),  # Q3 : Fléchir les jambes (environ 90°)
    (16, 92.5, 599.0),  # Q4 : Faux (ne pas plier les bras)
    (16, 92.5, 667.0),  # Q5 : Prise à pleine main (paume et doigts)
    (16, 428.0, 755.0), # Q6 : Garder la charge près du corps

    # ---------------- PAGE 18 (Validation FCGP) ----------------
    (17, 245.0, 555.0), # FC GP validée à 100% : OUI
]

def clean_inline_text(val):
    """Supprime les retours à la ligne (\\n, \\r) et tabulations, et normalise les espaces."""
    if val is None:
        return ""
    s = re.sub(r"[\r\n\t]+", " ", str(val))
    return re.sub(r"\s+", " ", s).strip()

def format_date(d_val):
    """Convertit dates Excel, YYYY-MM-DD ou nombres en format JJ/MM/AAAA."""
    if d_val is None:
        return ""
    if isinstance(d_val, datetime):
        return d_val.strftime("%d/%m/%Y")
    
    val_str = clean_inline_text(d_val)
    if re.match(r"^\d{5}(\.\d+)?$", val_str):
        try:
            excel_origin = datetime(1899, 12, 30)
            d = excel_origin + timedelta(days=float(val_str))
            return d.strftime("%d/%m/%Y")
        except:
            pass

    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", val_str)
    if m:
        return f"{int(m.group(3)):02d}/{int(m.group(2)):02d}/{m.group(1)}"
    
    m2 = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})", val_str)
    if m2:
        return f"{int(m2.group(1)):02d}/{int(m2.group(2)):02d}/{m2.group(3)}"

    return val_str

def parse_cell_ref(ref):
    """Convertit une référence de cellule ('A1', 'BC12') en (ligne 1-based, colonne 0-based)."""
    m = re.match(r"^([A-Za-z]+)(\d+)$", ref.strip())
    if not m:
        return 0, 0
    letters = m.group(1).upper()
    col = 0
    for char in letters:
        col = col * 26 + (ord(char) - ord("A") + 1)
    return int(m.group(2)), col - 1

def read_xlsx_native(path):
    """Lit un fichier .xlsx et résout les cellules fusionnées (sans dépendance externe)."""
    with zipfile.ZipFile(path) as z:
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            tree = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in tree.findall("{*}si"):
                text = "".join(t.text or "" for t in si.findall(".//{*}t"))
                strings.append(text)

        sheet_xml = z.read("xl/worksheets/sheet1.xml")
        tree = ET.fromstring(sheet_xml)
        
        grid = {}
        for row in tree.findall(".//{*}row"):
            r_idx = int(row.get("r", 0))
            if r_idx not in grid:
                grid[r_idx] = {}
            for c in row.findall("{*}c"):
                ref = c.get("r", "A1")
                r, col = parse_cell_ref(ref)
                t = c.get("t")
                v = c.find("{*}v")
                val = ""
                if v is not None and v.text is not None:
                    if t == "s":
                        s_idx = int(v.text)
                        val = strings[s_idx] if s_idx < len(strings) else ""
                    else:
                        val = v.text
                elif c.find("{*}is") is not None:
                    val = "".join(t.text or "" for t in c.findall(".//{*}t"))
                grid[r_idx][col] = val

        for mc in tree.findall(".//{*}mergeCell"):
            ref = mc.get("ref", "")
            if ":" in ref:
                start_ref, end_ref = ref.split(":")
                s_r, s_c = parse_cell_ref(start_ref)
                e_r, e_c = parse_cell_ref(end_ref)
                top_left_val = grid.get(s_r, {}).get(s_c, "")
                for r in range(s_r, e_r + 1):
                    if r not in grid:
                        grid[r] = {}
                    for col in range(s_c, e_c + 1):
                        grid[r][col] = top_left_val

        rows_data = []
        for r_idx in sorted(grid.keys()):
            cols = grid[r_idx]
            max_c = max(cols.keys()) if cols else -1
            row_vals = [cols.get(i, "") for i in range(max_c + 1)]
            if any(str(x).strip() for x in row_vals):
                rows_data.append(row_vals)

        return rows_data

def read_rows_from_file(input_path):
    """Lit les lignes d'un fichier Excel (.xlsx) ou CSV avec gestion des cellules fusionnées."""
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".xlsx":
        try:
            import openpyxl
            wb = openpyxl.load_workbook(input_path, data_only=True)
            ws = wb.active
            for rng in list(ws.merged_cells.ranges):
                min_col, min_row, max_col, max_row = rng.min_col, rng.min_row, rng.max_col, rng.max_row
                top_left_val = ws.cell(row=min_row, column=min_col).value
                ws.unmerge_cells(range_string=str(rng))
                for r in range(min_row, max_row + 1):
                    for c in range(min_col, max_col + 1):
                        ws.cell(row=r, column=c).value = top_left_val

            rows = []
            for r in ws.iter_rows(values_only=True):
                if any(r):
                    rows.append([str(c) if c is not None else "" for c in r])
            return rows
        except ImportError:
            return read_xlsx_native(input_path)

    encodings = ["utf-8-sig", "latin-1", "utf-8"]
    content = None
    for enc in encodings:
        try:
            with open(input_path, "r", encoding=enc) as f:
                content = f.read()
                break
        except UnicodeDecodeError:
            continue

    if content is None:
        raise ValueError(f"Impossible de lire le fichier '{input_path}'.")

    import csv
    sample = content[:2048]
    delimiter = ";" if sample.count(";") > sample.count(",") else ","
    return list(csv.reader(content.splitlines(), delimiter=delimiter))

def find_input_file():
    """Détecte automatiquement un fichier Excel ou CSV disponible."""
    excel_files = sorted(glob.glob("*.xlsx"))
    for ef in excel_files:
        if not os.path.basename(ef).startswith("~$"):
            return ef

    excel_sub = sorted(glob.glob("**/*.xlsx", recursive=True))
    for ef in excel_sub:
        if not os.path.basename(ef).startswith("~$") and "Dossiers_Generes" not in ef:
            return ef

    csv_candidates = [
        "liste_employes.csv",
        "Renouvellement Septembre/liste_employes.csv",
        "Renouvellement Juillet/liste_employes.csv",
    ]
    for c in csv_candidates:
        if os.path.exists(c):
            return c

    csv_matches = glob.glob("**/*employe*.csv", recursive=True)
    if csv_matches:
        return csv_matches[0]

    return None

def find_templates():
    """Détecte les modèles pré-corrigés disponibles."""
    templates = sorted(glob.glob("Intégration *.pdf"))
    if not templates:
        templates = sorted(glob.glob("**/Intégration *.pdf", recursive=True))
    if not templates:
        for fallback in ["Rev.pdf", "Renouvellement Septembre/Rev.pdf", "Renouvellement Juillet/Renouvellement.pdf"]:
            if os.path.exists(fallback):
                templates.append(fallback)
                break
    return templates

def parse_employees(rows):
    """Extrait la liste des employés et élimine les doublons causés par les fusions esthétiques."""
    if not rows:
        return []
    header = [str(h).strip().lower() for h in rows[0]]

    col_nom = next((i for i, h in enumerate(header) if "nom" in h or "employe" in h or "agent" in h), 0)
    col_date = next((i for i, h in enumerate(header) if "date" in h), 1 if len(header) > 1 else None)
    col_poste = next((i for i, h in enumerate(header) if "poste" in h or "fonction" in h or "role" in h), 2 if len(header) > 2 else None)

    employees = []
    seen = set()

    for r in rows[1:]:
        if not r or not any(str(x).strip() for x in r):
            continue
        nom = clean_inline_text(r[col_nom]).upper() if col_nom < len(r) else ""
        if not nom or nom.lower() == "nom":
            continue
        
        date_raw = r[col_date] if (col_date is not None and col_date < len(r)) else ""
        date = format_date(date_raw)
        raw_poste = r[col_poste] if (col_poste is not None and col_poste < len(r)) else ""
        poste = clean_inline_text(raw_poste).upper() if raw_poste else "EQUIPIER"

        emp_key = (nom, date, poste)
        if emp_key in seen:
            continue
        if not date and any(e["nom"] == nom for e in employees):
            continue

        seen.add(emp_key)
        employees.append({"nom": nom, "date": date, "poste": poste})

    return employees

def insert_auto_fit_text(page, point, text, max_width=None, base_fs=10, min_fs=6.0, fontname="helv", color=(0, 0, 0)):
    """Insère du texte en réduisant automatiquement la taille de police si le texte dépasse max_width."""
    fontsize = base_fs
    if max_width and text:
        while fontsize > min_fs and fitz.get_text_length(text, fontname=fontname, fontsize=fontsize) > max_width:
            fontsize -= 0.5
    page.insert_text(point, text, fontsize=fontsize, fontname=fontname, color=color)

def is_cross_circled(page, point):
    """Vérifie si une case (point) est entourée d'un rond rouge (signalant une faute/erreur du candidat)."""
    # 1. Vérification dans les annotations manuscrites (Ink)
    for a in page.annots():
        if a.type[0] == 15:  # Ink
            for s in a.vertices:
                xs = [pt[0] if isinstance(pt, (tuple, list)) else pt.x for pt in s]
                ys = [pt[1] if isinstance(pt, (tuple, list)) else pt.y for pt in s]
                w = max(xs) - min(xs)
                h = max(ys) - min(ys)
                if 8 <= w <= 65 and 8 <= h <= 65:
                    r = fitz.Rect(min(xs), min(ys), max(xs), max(ys))
                    if r.contains(point) or fitz.Rect(point.x - 6, point.y - 6, point.x + 6, point.y + 6).intersects(r):
                        return True

    # 2. Vérification dans les tracés vectoriels (drawings)
    for d in page.get_drawings():
        c = d.get("color")
        if c and len(c) == 3 and c[0] > 0.6 and c[1] < 0.4:
            r = fitz.Rect(d["rect"])
            if 8 <= r.width <= 65 and 8 <= r.height <= 65:
                if r.contains(point) or fitz.Rect(point.x - 6, point.y - 6, point.x + 6, point.y + 6).intersects(r):
                    return True

    return False

def generate_pdf(template_path, employee, output_path):
    """Génère le PDF personnalisé en appliquant les croix de correction et les annotations."""
    doc = fitz.open(template_path)
    nom = employee["nom"]
    date = employee["date"]
    poste = employee["poste"]

    # 0. Supprimer la couche interactive AcroForm (widgets vides)
    #    Empêche les lecteurs PDF (Adobe Acrobat, Chrome, etc.) de masquer les croix avec des cases blanches vides à partir de la page 5.
    doc.bake(annots=False, widgets=True)
    for p in doc:
        for w in list(p.widgets()):
            p.delete_widget(w)

    # 1. Application des numéros de l'Exercice 5 (Page 2)
    for pno, x, y, num in EXO_5_NUMBERS:
        if pno < len(doc):
            doc[pno].insert_text(
                fitz.Point(x, y),
                num,
                fontsize=11,
                fontname="hebo",
                color=(0, 0, 0)
            )

    # 2. Application des croix de correction calibrées (sauf si la case est entourée d'un rond rouge = faute)
    for pno, x, y in CORRECTION_CROSSES:
        if pno < len(doc):
            if not is_cross_circled(doc[pno], fitz.Point(x, y)):
                doc[pno].insert_text(
                    fitz.Point(x, y),
                    "X",
                    fontsize=10.5,
                    fontname="hebo",
                    color=(0, 0, 0)
                )

    # 3. Insertion des données personnalisées de l'employé avec auto-fit
    for item in TEXT_FIELDS:
        page_idx, x, y, key = item[0], item[1], item[2], item[3]
        max_w = item[4] if len(item) > 4 else None
        if page_idx < len(doc):
            val = nom if key == "nom" else (date if key == "date" else poste)
            if val:
                insert_auto_fit_text(doc[page_idx], fitz.Point(x, y), val, max_width=max_w)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    doc.close()

def find_ab_template():
    """Détecte si un modèle Évaluation AB.pdf est présent."""
    matches = glob.glob("*AB*.pdf") + glob.glob("**/*AB*.pdf", recursive=True)
    return matches[0] if matches else None

def generate_evaluation_ab(template_path, employee, output_path):
    """Génère la fiche Évaluation AB personnalisée."""
    doc = fitz.open(template_path)
    nom = employee["nom"]
    date = employee["date"]

    valid_checks = {
        0: ["Case à cocher17", "Case à cocher16", "Case à cocher15", "Case à cocher14", "Case à cocher13", "Case à cocher12"],
        1: ["Case à cocher18", "Case à cocher19", "Case à cocher35", "Case à cocher34", "Case à cocher33"],
        3: ["Case à cocher8", "Case à cocher9", "Case à cocher10", "Case à cocher11", "Case à cocher12"],
        4: ["Case à cocher1", "Case à cocher2"]
    }

    for p_no in range(len(doc)):
        p = doc[p_no]
        checks_to_set = valid_checks.get(p_no, [])
        for w in p.widgets():
            if w.field_name in checks_to_set:
                w.field_value = "Oui"
                w.update()
            elif p_no == 2:  # Page 3
                if "Date" in w.field_name and date:
                    w.field_value = date
                    w.update()
                elif "Texte39" in w.field_name or "Texte38" in w.field_name:
                    if w.rect.x0 > 300 and nom:
                        w.field_value = nom
                        w.update()
            elif p_no == 4:  # Page 5
                if "Date" in w.field_name and date:
                    w.field_value = date
                    w.update()
                elif "Texte1" in w.field_name or "Texte23" in w.field_name:
                    if w.rect.x0 > 300 and nom:
                        w.field_value = nom
                        w.update()

    doc.bake(annots=False, widgets=True)
    for p in doc:
        for w in list(p.widgets()):
            p.delete_widget(w)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    doc.close()

def main():
    parser = argparse.ArgumentParser(description="Générateur de fiches PDF avec tirage aléatoire, cellules fusionnées et correction automatique calibrée.")
    parser.add_argument("input_file", nargs="?", default=None, help="Chemin vers le fichier Excel (.xlsx) ou CSV (optionnel).")
    parser.add_argument("--outdir", default="Dossiers_Generes", help="Dossier de destination (défaut : Dossiers_Generes).")
    args = parser.parse_args()

    input_path = args.input_file or find_input_file()
    if not input_path or not os.path.exists(input_path):
        print("❌ Erreur : Aucun fichier d'entrée (Excel .xlsx ou CSV) trouvé.")
        print("   Déposez simplement un fichier Excel (.xlsx) dans ce dossier et relancez le script.")
        sys.exit(1)

    templates = find_templates()
    if not templates:
        print("❌ Erreur : Aucun modèle PDF ('Intégration *.pdf') trouvé.")
        sys.exit(1)

    ab_template = find_ab_template()

    print("=" * 60)
    print("   AUTOMATISATION DES FICHES D'ÉVALUATION")
    print("=" * 60)
    print(f"📊 Fichier source : {input_path}")
    print(f"📄 Modèles disponibles ({len(templates)}) : {', '.join(os.path.basename(t) for t in templates)}")
    if ab_template:
        print(f"📋 Fiche complémentaire : {os.path.basename(ab_template)}")
    print(f"📁 Dossier de sortie : {args.outdir}")
    print(f"🎲 Mode d'attribution : Tirage aléatoire parmi les {len(templates)} modèles")
    print(f"🔢 Exercice 5 (Page 2) : Numérotation des 7 étapes (4, 1, 7, 5, 2, 6, 3)")
    print(f"🎯 Calibrage précis : Correction des coordonnées US Letter (hauteur 792 pt)")
    print(f"✅ Remplissage des corrections officielles : Actif ({len(CORRECTION_CROSSES)} croix)")
    print("-" * 60)

    rows = read_rows_from_file(input_path)
    employees = parse_employees(rows)
    print(f"👥 {len(employees)} employé(s) chargé(s).\n")

    start_time = time.time()
    for emp in employees:
        chosen_template = random.choice(templates)
        tmpl_name = os.path.basename(chosen_template)

        out_filename = f"Fiche_{emp['nom']}.pdf"
        out_path = os.path.join(args.outdir, out_filename)

        generate_pdf(chosen_template, emp, out_path)
        if ab_template:
            ab_filename = f"Evaluation_AB_{emp['nom']}.pdf"
            ab_path = os.path.join(args.outdir, ab_filename)
            generate_evaluation_ab(ab_template, emp, ab_path)
            print(f"  🎲 {emp['nom']:<26} -> {tmpl_name} + {os.path.basename(ab_template)}")
        else:
            print(f"  🎲 {emp['nom']:<26} -> {tmpl_name} -> {out_filename}")

    elapsed = time.time() - start_time
    print("-" * 60)
    print(f"✨ Terminé avec succès ! {len(employees)} fiches générées en {elapsed:.2f}s.")
    print(f"📂 Retrouvez vos fichiers dans le dossier : {os.path.abspath(args.outdir)}")
    print("=" * 60)

if __name__ == "__main__":
    main()
