# SCG-Ré — Tableau de Bord Dynamique de Production

Ce tableau de bord interactif autonome a été conçu pour suivre dynamiquement la production de souscription (Cessions Commerciales FAC) de la SCG-Ré, sans nécessiter d'installation ni de serveur.

---

## 🚀 Comment lancer le tableau de bord

1. Double-cliquez simplement sur **`Lancer_Dashboard.bat`** (ou ouvrez directement **`index.html`** dans **Microsoft Edge**, **Google Chrome** ou **Firefox**).
2. Le tableau de bord s'ouvre **instantanément** avec l'ensemble des **467 affaires** préchargées et nettoyées, en **pleine largeur d'écran**.

---

## 💎 Nouveautés & Évolutions Récentes

* **Disposition Pleine Largeur (100% Full Width)** : Affichage fluide sur toute la largeur d'écran pour un confort de lecture optimal.
* **Séparateurs d'espaces sur tous les montants** : Formatage financier standardisé (ex: `1 853 088 489 FCFA`) sur l'ensemble des indicateurs, tableaux et infobulles.
* **Étiquettes de valeurs sur les graphiques** : Les montants s'affichent directement au-dessus / à l'extrémité de chaque barre (`chartjs-plugin-datalabels`).
* **5 Indicateurs Permanents Dynamiques** :
  1. *Capitaux totaux (100%)*
  2. *Engagement SCG-Ré*
  3. *Prime SCG-Ré (Brute)*
  4. *Commission*
  5. *Prime nette*
  *(Les KPI s'adaptent dynamiquement au contexte de l'onglet actif : production souscrite globale vs production signée avec bordereaux reçus).*
* **En-tête épuré** : Conservation du seul filtre **Branche** (avec réinitialisation et compteur en direct).
* **Standardisation Poussée des Données** :
  - **Apporteurs** : Correction des fautes de frappe et variantes (`GLOBUS RE`, `PREVEX RE`, `ATLAS REINSURANCE CONSULTANTS`, `SB RE`, `KLAPTON AFRICA`, `SUNU RE`).
  - **Cédantes** : Regroupement systématique par **Groupe d'assurance** (`SUNU ASSURANCES`, `GNA`, `NSIA`, `CORIS ASSURANCES`, `SANLAM ALLIANZ`, `ATLANTIQUE ASSURANCES`, `ACTIVA`, `AMSA`, `VISTA ASSURANCES`, etc.) et qualification des co-assurances en `CO-ASSURANCE / APÉRITEUR`.

---

## 📊 Les 3 Onglets du Tableau de Bord

1. **Vue Prime souscrite** :
   - Primes brutes SCG-Ré ventilées par **Branche**, **Pays**, **Cédante (Groupe)** et **Courtier / Apporteur**.
   - Tout est classé par **ordre strictement décroissant**.
   - Bascule instantanée Graphique Top 10 / Tableau complet avec barres de progression pour les Cédantes et Courtiers.
2. **Vue Prime signée** :
   - Concerne les affaires fermes dont les bordereaux ont été effectivement reçus par la SCG-Ré.
   - Primes signées ventilées par **Branche**, **Pays**, **Cédante (Groupe)** et **Courtier / Apporteur** par **ordre strictement décroissant**.
   - Bascule instantanée Graphique Top 10 / Tableau complet.
3. **Explorateur de Données & Export** :
   - Recherche textuelle instantanée, pagination dynamique (25, 50, 100, Tout), tri multi-colonnes.
   - Bouton d'export Excel direct du portefeuille filtré.

---

## 📂 Mise à jour des données (Glisser-Déposer / Drag & Drop)

* Cliquez sur le bouton en haut à droite : **"Importer Nouveau Fichier Excel"**.
* Glissez-déposez votre nouveau classeur `.xlsx`.
* Le moteur JavaScript intégré traite le fichier en moins de 0.2 seconde, applique automatiquement la standardisation (branches, pays, groupes de cédantes et courtiers) et rafraîchit l'ensemble des indicateurs et graphiques.

---

## 📁 Fichiers inclus dans le dossier

* **`index.html`** : L'application web autonome complète du tableau de bord.
* **`generate_dashboard.py`** : Script Python de génération et régénération du tableau de bord.
* **`cleaned_records.json`** : Données de production 100% standardisées avec journal de nettoyage.
* **`Lancer_Dashboard.bat`** : Raccourci pour ouvrir le tableau de bord en un clic.
* **`Classeur1.xlsx`** : Le fichier source d'origine.
