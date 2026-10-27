#!/usr/bin/env bash
# Script de lancement 1-clic pour Linux / macOS
set -e

cd "$(dirname "$0")"

echo "============================================================"
echo "       GÉNÉRATION DES FICHES D'ÉVALUATION (LINUX)"
echo "============================================================"
echo ""

# 1. Utilisation prioritaire de uv si disponible
if command -v uv &> /dev/null; then
    echo "[OK] Utilisation du moteur uv..."
    uv run --with pymupdf python3 Main.py "$@"
# 2. Utilisation de python3 si PyMuPDF est déjà installé
elif python3 -c "import fitz" &> /dev/null; then
    echo "[OK] Utilisation de Python 3 système avec fitz..."
    python3 Main.py "$@"
# 3. Création automatique d'un environnement virtuel isolé (compatible PEP 668)
elif command -v python3 &> /dev/null; then
    echo "[INFO] Configuration de l'environnement virtuel local..."
    if [ ! -d ".venv_linux" ]; then
        python3 -m venv .venv_linux
    fi
    echo "[INFO] Installation de pymupdf dans l'environnement virtuel..."
    .venv_linux/bin/pip install --quiet pymupdf
    .venv_linux/bin/python Main.py "$@"
else
    echo "[ERREUR] Aucun interpréteur Python 3 ou uv trouvé sur cette machine."
    echo "Installez Python 3 avec : sudo apt install python3 python3-venv (Debian/Ubuntu)"
    exit 1
fi

echo ""
echo "============================================================"
echo "  ✨ Fiches prêtes dans le dossier : Dossiers_Generes"
echo "============================================================"

# Ouvrir le dossier si un gestionnaire de fenêtres est actif
if command -v xdg-open &> /dev/null && [ -n "$DISPLAY" ]; then
    xdg-open Dossiers_Generes &> /dev/null &
fi
