#!/usr/bin/env python3
"""
Remplissage automatique des fiches d'évaluation (Version Septembre).
Redirige vers le script racine ou fonctionne en autonome avec PyMuPDF vectoriel.
"""

import os
import sys

# Si le Main.py racine existe, on l'utilise directement
ROOT_MAIN = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Main.py"))
if os.path.exists(ROOT_MAIN):
    # Change working directory vers la racine pour accéder aux Intégration *.pdf
    os.chdir(os.path.dirname(ROOT_MAIN))
    with open(ROOT_MAIN, "rb") as f:
        code = compile(f.read(), ROOT_MAIN, 'exec')
        exec(code)
else:
    print("Erreur : Main.py racine introuvable.")
