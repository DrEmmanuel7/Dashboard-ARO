# SCG-Ré — Tableau de Bord Dynamique de Production

Ce tableau de bord interactif autonome a été conçu pour suivre dynamiquement la production de souscription (Cessions Commerciales FAC) de la SCG-Ré, sans nécessiter d'installation ni de serveur.

---

## 🚀 Comment lancer le tableau de bord

1. Double-cliquez simplement sur **`Lancer_Dashboard.bat`** (ou ouvrez directement **`index.html`** dans **Microsoft Edge**, **Google Chrome** ou **Firefox**).
2. Le tableau de bord s'ouvre **instantanément** avec l'ensemble des **467 affaires** préchargées et nettoyées, en **pleine largeur d'écran**.

---

## 💎 Nouveautés & Évolutions Récentes

* **Refonte Stylistique « Executive Segmented » des Boutons & Onglets** :
  - Barre d'onglets moderne en conteneur segmenté avec codes couleurs distinctifs : Bleu royal pour la *Vue Souscrite*, Vert émeraude pour la *Vue Signée*, et Ardoise sombre pour l'*Explorateur de Données*.
  - Badges intégrés dynamiques affichant en temps réel le nombre d'affaires ou de bordereaux filtrés.
  - Boutons d'action haut de gamme (import, export, impression, réinitialisation) avec micro-interactions et ombres douces.
* **Étanchéité Absolue de l'Onglet « Prime signée »** :
  - Exclusion totale de toute donnée ou dénominateur de prime souscrite sur l'onglet Prime signée.
  - Les 5 KPI et la barre d'état du filtre sont calculés et libellés exclusivement sur le périmètre des **bordereaux reçus** (*« Affichage de X bordereaux reçus sur N_signés »*).
* **Double Vue (Graphique Top / Tableau Complet) sur TOUS les Blocs** :
  - Bascule instantanée disponible sur les 4 dimensions : **Branche**, **Pays**, **Cédante** et **Courtier / Apporteur**.
  - Tableaux complets riches et homogènes intégrant : **Rang** (badges dorés/argentés/bronzés pour le Top 3), **Libellé**, **Prime SCG-Ré**, **Nombre d'affaires / de bordereaux** et **Part (%)** avec jauge de progression visuelle colorée.
* **Tableau Dédié « Top 10 des Affaires Signées »** :
  - Situé en pleine largeur sous les 4 blocs dans l'onglet Prime signée.
  - Présentation détaillée des 10 plus importantes affaires fermes avec bordereaux reçus : *Affaire*, *Cédante*, *Courtier*, *Capitaux 100%*, *Prime 100%*, *Taux de cession*, *Engagement SCG-Ré*, et *Prime SCG-Ré* (mise en valeur en vert émeraude).
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

## 📂 Mise à jour et Sauvegarde des Données (Persistance Automatique)

* **Importation Glisser-Déposer instantanée** :
  1. Cliquez sur le bouton en haut à droite : **"Importer Nouveau Fichier Excel"**.
  2. Glissez-déposez votre nouveau classeur `.xlsx`.
  3. Le moteur intégré traite le fichier immédiatement, standardise les données et actualise tous les graphiques.
* **Mémorisation Permanente & Par Défaut (Nouveau !)** :
  - La nouvelle base importée est désormais **automatiquement enregistrée** dans le stockage persistant de votre navigateur (`IndexedDB` + `localStorage`).
  - **Au redémarrage du PC, fermeture du navigateur ou en recliquant sur le lien / `Lancer_Dashboard.bat`**, vos nouvelles données restent enregistrées et sont **utilisées par défaut**.
  - Un badge vert dans l'en-tête indique la base active, la date d'enregistrement et le nombre d'affaires.
* **Bouton « Rétablir base initiale »** :
  - Vous pouvez revenir à tout moment à la base de données d'origine pré-intégrée (498 affaires) en un seul clic sur le bouton de réinitialisation.
* **Bouton « Télécharger HTML à jour »** :
  - Permet d'exporter et de télécharger un fichier `index.html` autonome autonome contenant directement votre nouvelle base en dur pour le partager ou remplacer le fichier sur votre disque.

---

## 📁 Fichiers inclus dans le dossier

* **`index.html`** : L'application web autonome complète du tableau de bord.
* **`generate_dashboard.py`** : Script Python de génération et régénération du tableau de bord.
* **`cleaned_records.json`** : Données de production 100% standardisées avec journal de nettoyage.
* **`Lancer_Dashboard.bat`** : Raccourci pour ouvrir le tableau de bord en un clic.
* **`Classeur1.xlsx`** : Le fichier source d'origine.
