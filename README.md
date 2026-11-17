# SCG-Ré — Tableau de Bord Dynamique de Production FAC ARO

Ce tableau de bord interactif autonome a été conçu pour suivre dynamiquement la production de souscription (Cessions Commerciales FAC) de la **Société Commerciale Gabonaise de Réassurance (SCG-Ré)**, sans nécessiter d'installation ni de serveur.

---

## 🚀 Comment lancer le tableau de bord

1. Double-cliquez simplement sur **`Lancer_Dashboard.bat`** (ou ouvrez directement **`index.html`** dans Microsoft Edge, Google Chrome ou Firefox).
2. Le tableau de bord s'ouvre instantanément en **pleine largeur d'écran** avec l'ensemble des données préchargées et nettoyées.

---

## 💎 Nouveautés & Évolutions Majeures

### 1. 🚫 Nouvel Onglet Dédié : « Affaires Sans Suite » (Exercice 2026)
* **Périmètre & Détection Automatique** :
  - Détection automatique via la colonne `COMMENTAIRES` (détection des mots-clés *"sans suite"* et *"classé sans suite"*).
  - Présentation exclusive de la déperdition sur la campagne 2026 sans fausser les chiffres souscrits.
* **4 Indicateurs Clés Dédiés** :
  1. *Nombre d'affaires sans suite* (avec taux de déperdition % par rapport au total des affaires).
  2. *Primes 100% manquées* (assiette brute totale de primes non concrétisées en FCFA).
  3. *Capitaux 100% abandonnés* (valeurs assurées totales non retenues en FCFA).
  4. *Prime 100% moyenne* par affaire sans suite.
* **4 Blocs de Ventilation (Double Vue Graphique Top / Tableau Complet)** :
  - Classement décroissant par Primes 100% sur les 4 dimensions : **Branche**, **Pays**, **Cédante** et **Courtier / Apporteur**.
* **Tableau Détaillé & Export Excel Spécifique** :
  - Liste intégrale des affaires sans suite avec motifs et commentaires.
  - Recherche instantanée dédiée et bouton d'export Excel direct (`SCG_Re_Affaires_Sans_Suite_2026.xlsx`).

---

### 2. ⚖️ Analyse Comparative 2025 vs 2026 (Onglet « Prime signée »)
* **Sélecteur Segmenté [ 2026 | Comparatif ]** :
  - Situé dans la bannière de l'onglet Prime signée.
  - Bascule instantanée entre la vue standard 2026 et l'analyse comparative avec l'exercice 2025.
* **Sous-Lignes KPI Comparatives Dynamiques** :
  - Sous chacun des 5 KPI (Capitaux signés, Engagement SCG-Ré, Prime signée, Commission, Prime nette), affichage direct du montant correspondant de 2025 (`2025 : X XXX XXX FCFA`).
  - Badge de variation coloré avec signe et pourcentage (`+X.X%` en vert émeraude, `-X.X%` en rose/rouge, `0.0%` en ardoise).
* **Graphiques en Double Barre (Grouped Bars)** :
  - Visualisation comparative côte à côte : Barre 2026 (vert émeraude) vs Barre 2025 (ardoise/acier).
  - Étiquettes de valeurs dynamiques au bout des barres et infobulles détaillées.
* **Tableaux Comparatifs Complets** :
  - Colonnes comparatives détaillées : *Rang*, *Libellé*, *Prime SCG 2026*, *Prime SCG 2025*, *Évolution (FCFA)*, *Variation (%)* et *Bordereaux (26 vs 25)*.

---

### 3. 📂 Explorateur de Données Enrichi
* **Sélecteur d'Exercice Net [ 2026 | 2025 ]** :
  - Permet d'explorer et d'exporter alternativement le portefeuille 2026 (en cours) ou le portefeuille 2025 (référentiel historique).
* **Badges de Statut Visuels** :
  - Badges distinctifs : *SIGNÉE* (vert émeraude), *SANS SUITE* (rose), *EN ATTENTE* (ambre).
* **Recherche & Export multi-colonnes** instantané vers Excel (.xlsx).

---

### 4. 🔄 Moteur de Standardisation & Fichiers
* **Standardisation Homogène 2025 et 2026** :
  - Application stricte des mêmes règles de nettoyage pour les Cédantes (Groupes consolidés), Apporteurs/Courtiers, Branches et Pays sur les deux exercices.
  - Intégration figée du fichier `Production 2025.xlsx` comme benchmark historique de comparaison.
* **Actualisation Automatique 2026** :
  - Double-cliquez sur `Mettre_A_Jour_Donnees_ARO.bat` pour actualiser automatiquement les données avec le dernier fichier de production 2026.
  - Ou utilisez le bouton *"Importer Excel"* directement dans l'interface web (glisser-déposer).

---

## 📁 Fichiers inclus dans le dossier

* **`index.html`** : L'application web autonome complète du tableau de bord (aucun serveur requis).
* **`generate_dashboard.py`** : Script Python de compilation du tableau de bord.
* **`update_data.py`** : Script de détection, nettoyage et standardisation multi-exercices (2025 et 2026).
* **`Mettre_A_Jour_Donnees_ARO.bat`** : Raccourci pour lancer la mise à jour des données en un clic.
* **`Lancer_Dashboard.bat`** : Raccourci pour ouvrir le tableau de bord dans le navigateur.
* **`cleaned_records.json`** : Données de production standardisées 2025 et 2026.
* **`Production 2025.xlsx`** : Référentiel historique de production 2025.
* **`Classeur1.xlsx`** : Fichier source de base 2026.
