@echo off
chcp 65001 >nul
title Générateur de Fiches d'Évaluation
cd /d "%~dp0"

echo ============================================================
echo       GÉNÉRATION DES FICHES D'ÉVALUATION (PORTABLE)
echo ============================================================
echo.

if exist "python_portable\python.exe" (
    echo [OK] Utilisation de l'environnement Python portable embarqué...
    "python_portable\python.exe" Main.py %*
) else if exist "uv.exe" (
    echo [OK] Utilisation du moteur portable uv.exe...
    "uv.exe" run --with pymupdf python Main.py %*
) else (
    where python >nul 2>nul
    if %errorlevel% equ 0 (
        echo [INFO] Utilisation du Python système...
        python -c "import fitz" >nul 2>nul
        if %errorlevel% neq 0 (
            echo Installation de la bibliothèque PDF pymupdf...
            python -m pip install --quiet pymupdf
        )
        python Main.py %*
    ) else (
        echo [ERREUR] Aucun moteur Python trouvé.
        echo Assurez-vous que le dossier 'python_portable' ou 'uv.exe' est bien présent.
        pause
        exit /b 1
    )
)

echo.
echo ============================================================
echo   ✨ Fiches prêtes dans le dossier : Dossiers_Generes
echo ============================================================
echo.
if exist "Dossiers_Generes" (
    echo Ouverture automatique du dossier...
    start "" "Dossiers_Generes"
)

echo.
echo Appuyez sur une touche pour fermer cette fenêtre...
pause >nul
