# Automatisation des Fiches d'Évaluation & Intégration

Projet d'automatisation et de personnalisation des dossiers d'évaluation (FCH, FCS, FCAPH, FCGP) et des fiches d'évaluation complémentaire (Évaluation AB).

---

## 🚀 Utilisation en 1 Clic

### Sous Windows
* Double-cliquez simplement sur **`LANCER_GENERATION.bat`**.
* Le script utilise l'environnement portable autonome (`python_portable/` ou `uv.exe`).
* Aucune installation de Python requise, aucun droit administrateur nécessaire, fonctionnement 100% hors-ligne.
* Le dossier **`Dossiers_Generes`** s'ouvre automatiquement à la fin de la génération.

### Sous Linux / macOS
* Ouvrez un terminal dans le dossier et lancez :
  ```bash
  ./lancer.sh
  ```
* Le script détecte `uv` ou `python3`, configure un environnement virtuel local si nécessaire (compatible Debian/Ubuntu PEP 668) et génère les dossiers.

---

## 📁 Données d'entrée

* **Fichier source** : `liste_employes.xlsx` (ou `liste_employes.csv`).
* **Colonnes supportées** :
  * `nom` : Nom de l'employé (gère les fusions de cellules, sauts de ligne internes et auto-réduction de police pour les noms longs).
  * `date` : Date d'évaluation (formatée automatiquement en `JJ/MM/AAAA`).
  * `poste` : Fonction (ex. *Equipier*, *Manager*, etc.).

---

## 🎯 Fonctionnalités du Programme (`Main.py`)

1. **Tirage aléatoire équitable** parmi les 3 modèles pré-corrigés officiels (`Intégration 001.pdf`, `Intégration 002.pdf`, `Intégration 003.pdf`).
2. **Prise en charge de la fiche complémentaire** : si `Évaluation AB.pdf` est présent, une évaluation AB personnalisée est automatiquement produite pour chaque employé.
3. **Respect du corrigé officiel et des fautes réelles** :
   * Détection géométrique des annotations manuscrites de correction (`is_cross_circled`).
   * Les cases entourées d'un rond rouge (signalant une faute) restent vides pour préserver la cohérence avec les annotations du formateur.
   * Les sous-parties partielles (ex. `1/3`) sont fidèlement respectées.
4. **Exercice 5 (Page 2)** : Remise dans l'ordre chronologique des 7 étapes (`4, 1, 7, 5, 2, 6, 3`).
5. **Calibrage US Letter (Pages 10 à 13)** : Correction dynamique de la hauteur pour centrer parfaitement les croix dans les cases.
6. **Aplatissement AcroForm universel** : Suppression des masques interactifs vides pour garantir une visibilité parfaite sous Adobe Acrobat, Chrome et Edge.

---

## 🔒 Confidentialité & Données

* Toutes les données nominatives contenues dans les fichiers d'exemple ont été anonymisées.
* Les dossiers générés (`Dossiers_Generes/`) et les binaires volumineux sont automatiquement exclus du suivi de version via `.gitignore`.
