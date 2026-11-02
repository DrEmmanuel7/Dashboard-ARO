import json
import os
import re
import base64

print("Starting dashboard generator...")

with open('cleaned_records.json', 'r', encoding='utf-8') as f:
    payload = json.load(f)

records_json = json.dumps(payload['records'], ensure_ascii=False)
log_json = json.dumps(payload.get('cleaning_log', []), ensure_ascii=False)

# Read logo if exists for base64 embed
logo_b64 = ""
if os.path.exists('logo_scg_re.png'):
    with open('logo_scg_re.png', 'rb') as f:
        logo_b64 = "data:image/png;base64," + base64.b64encode(f.read()).decode('utf-8')

html_template = """<!DOCTYPE html>
<html lang="fr" class="h-full bg-slate-50">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SCG-Ré | Tableau de Bord ARO</title>
  
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      darkMode: 'class',
      theme: {
        extend: {
          colors: {
            brand: {
              50: '#f0f7ff',
              100: '#e0effe',
              200: '#bae0fd',
              500: '#0284c7',
              600: '#0369a1',
              700: '#075985',
              800: '#0c4a6e',
              900: '#082f49',
              950: '#041e33',
            },
            emerald: {
              500: '#10b981',
              600: '#059669',
              700: '#047857',
            }
          }
        }
      }
    }
  </script>

  <!-- Chart.js & Plugins -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
  <!-- SheetJS (xlsx) for instant client-side Excel parse & export -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    body {
      font-family: 'Plus Jakarta Sans', sans-serif;
    }
    .custom-scrollbar::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    .custom-scrollbar::-webkit-scrollbar-track {
      background: #f1f5f9;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 4px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover {
      background: #94a3b8;
    }
    @media print {
      .no-print { display: none !important; }
      body { background: white !important; }
      .shadow-sm, .shadow-md, .shadow-lg { box-shadow: none !important; border: 1px solid #e2e8f0; }
    }
  </style>
</head>
<body class="h-full flex flex-col text-slate-800 antialiased selection:bg-brand-500 selection:text-white">

  <!-- TOP HEADER (PLEINE LARGEUR AVEC LOGO OFFICIEL SCG-RÉ) -->
  <header class="bg-gradient-to-r from-brand-950 via-brand-900 to-slate-900 text-white shadow-lg sticky top-0 z-40 border-b border-brand-800/50">
    <div class="w-full px-4 sm:px-6 lg:px-8 py-3 flex flex-wrap items-center justify-between gap-4">
      
      <!-- Brand & Title -->
      <div class="flex items-center space-x-3.5">
        <div class="h-12 px-2.5 bg-white rounded-xl shadow-md ring-2 ring-white/20 flex items-center justify-center">
          <img src="__LOGO_SRC__" alt="SCG-Ré" class="h-9 w-auto object-contain" />
        </div>
        <div>
          <div class="flex items-center gap-2">
            <h1 class="text-xl font-extrabold tracking-tight text-white flex items-center gap-2.5">
              Tableau de Bord ARO
              <span class="text-xs font-semibold uppercase tracking-wider bg-emerald-500/20 text-emerald-300 px-2.5 py-0.5 rounded-full border border-emerald-500/30">Cessions FAC</span>
            </h1>
          </div>
          <p class="text-xs text-brand-200/80 font-medium flex items-center gap-2 mt-0.5">
            <span>SCG-Ré</span>
            <span>•</span>
            <span id="header-records-count">498 affaires chargées</span>
            <span>•</span>
            <span class="text-emerald-400 font-semibold flex items-center gap-1">
              <span class="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span> Données Standardisées
            </span>
          </p>
        </div>
      </div>

      <!-- Action Buttons -->
      <div class="flex items-center gap-2.5 no-print">
        <button id="btn-open-upload" class="inline-flex items-center gap-1.5 px-3 py-1.5 bg-brand-700/70 hover:bg-brand-600 text-brand-50 rounded-lg text-xs font-semibold shadow-sm transition-all border border-brand-500/30 backdrop-blur-sm">
          <svg class="w-4 h-4 text-brand-300" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>
          Importer Nouveau Fichier Excel
        </button>

        <button id="btn-export-cleaned" class="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold shadow-sm transition-all">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
          Exporter Excel Propre
        </button>

        <button onclick="window.print()" class="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-semibold border border-slate-700 transition-all">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"/></svg>
          Imprimer / PDF
        </button>
      </div>

    </div>
  </header>

  <!-- STICKY FILTER BAR (PLEINE LARGEUR - FILTRE UNIQUE BRANCHE) -->
  <section class="bg-white border-b border-slate-200 shadow-sm sticky top-[65px] z-30 no-print">
    <div class="w-full px-4 sm:px-6 lg:px-8 py-3">
      <div class="flex flex-wrap items-center justify-between gap-4">
        
        <!-- Filter Unique: Branche -->
        <div class="flex items-center gap-3 flex-1 max-w-md">
          <label for="filter-branche" class="text-xs font-bold uppercase tracking-wider text-slate-600 whitespace-nowrap flex items-center gap-1.5">
            <svg class="w-4 h-4 text-brand-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/></svg>
            Filtrer par Branche :
          </label>
          <select id="filter-branche" class="w-full bg-slate-50 border border-slate-300 text-slate-700 text-xs font-medium rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition-colors shadow-xs">
            <option value="ALL">Toutes les branches</option>
          </select>
          <button id="btn-reset-filters" title="Réinitialiser le filtre" class="inline-flex items-center justify-center p-1.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg border border-slate-300 transition-colors">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
          </button>
        </div>

        <!-- Active filter status pill & Monnaie -->
        <div class="flex items-center gap-4 text-xs text-slate-500">
          <div id="filter-status-text" class="font-medium text-slate-600 bg-slate-100 px-3 py-1 rounded-full border border-slate-200">
            Affichage de <span class="font-bold text-brand-600" id="filtered-count">0</span> affaires sur <span class="font-bold" id="total-count-badge">0</span>
          </div>
          <div class="text-[11px] text-slate-400 font-medium">
            Monnaie de référence : <strong class="text-slate-700 font-semibold">FCFA (XOF)</strong>
          </div>
        </div>

      </div>
    </div>
  </section>

  <!-- NAVIGATION TABS (PLEINE LARGEUR - 3 ONGLETS) -->
  <nav class="bg-white border-b border-slate-200 no-print">
    <div class="w-full px-4 sm:px-6 lg:px-8">
      <div class="flex space-x-2 sm:space-x-4 overflow-x-auto custom-scrollbar py-2 text-xs font-semibold" id="tab-nav-container">
        
        <button data-tab="tab-souscrite" class="tab-btn active px-4 py-2.5 rounded-lg bg-brand-50 text-brand-700 border border-brand-200/80 shadow-xs flex items-center gap-2 whitespace-nowrap transition-all">
          <svg class="w-4 h-4 text-brand-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/></svg>
          1. Vue Prime souscrite (Valeurs brutes SCG-Ré)
        </button>

        <button data-tab="tab-signee" class="tab-btn px-4 py-2.5 rounded-lg text-slate-600 hover:text-slate-900 hover:bg-slate-100 flex items-center gap-2 whitespace-nowrap transition-all">
          <svg class="w-4 h-4 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
          2. Vue Prime signée (Bordereaux reçus)
        </button>

        <button data-tab="tab-data" class="tab-btn px-4 py-2.5 rounded-lg text-slate-600 hover:text-slate-900 hover:bg-slate-100 flex items-center gap-2 whitespace-nowrap transition-all">
          <svg class="w-4 h-4 text-slate-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h18M3 14h18m-9-4v8m-7 4h14a2 2 0 002-2V6a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>
          3. Explorateur de Données & Export
        </button>

      </div>
    </div>
  </nav>

  <!-- MAIN CONTENT CONTAINER (PLEINE LARGEUR) -->
  <main class="flex-1 w-full px-4 sm:px-6 lg:px-8 py-6 space-y-6">

    <!-- 5 PERMANENT KPI CARDS RIBBON -->
    <section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
      
      <!-- KPI 1: Capitaux Totaux 100% -->
      <div class="bg-white rounded-xl border border-slate-200/90 p-4 shadow-xs relative overflow-hidden group hover:border-indigo-300 transition-all">
        <div class="flex items-center justify-between text-slate-500 mb-1">
          <span class="text-xs font-semibold text-slate-600" id="kpi-label-capitaux-100">Capitaux Totaux (100%)</span>
          <span class="p-1 rounded bg-indigo-50 text-indigo-600">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg>
          </span>
        </div>
        <div class="text-xl font-extrabold text-slate-900 tracking-tight" id="kpi-capitaux-100">0 FCFA</div>
        <div class="text-[11px] text-slate-500 mt-1 flex items-center justify-between">
          <span>Assiette globale :</span>
          <span class="font-bold text-indigo-600" id="kpi-sub-capitaux-100">100% Marché</span>
        </div>
        <div class="absolute bottom-0 left-0 right-0 h-1 bg-indigo-500"></div>
      </div>

      <!-- KPI 2: Engagement SCG-Ré -->
      <div class="bg-white rounded-xl border border-slate-200/90 p-4 shadow-xs relative overflow-hidden group hover:border-purple-300 transition-all">
        <div class="flex items-center justify-between text-slate-500 mb-1">
          <span class="text-xs font-semibold text-slate-600" id="kpi-label-capitaux-scg">Engagement SCG-Ré</span>
          <span class="p-1 rounded bg-purple-50 text-purple-600">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/></svg>
          </span>
        </div>
        <div class="text-xl font-extrabold text-purple-900 tracking-tight" id="kpi-capitaux-scg">0 FCFA</div>
        <div class="text-[11px] text-slate-500 mt-1 flex items-center justify-between">
          <span>Part des capitaux :</span>
          <span class="font-bold text-purple-600" id="kpi-part-capitaux-scg">0%</span>
        </div>
        <div class="absolute bottom-0 left-0 right-0 h-1 bg-purple-500"></div>
      </div>

      <!-- KPI 3: Prime SCG-Ré -->
      <div class="bg-white rounded-xl border border-slate-200/90 p-4 shadow-xs relative overflow-hidden group hover:border-brand-300 transition-all">
        <div class="flex items-center justify-between text-slate-500 mb-1">
          <span class="text-xs font-semibold text-slate-600" id="kpi-label-primes-scg">Prime SCG-Ré</span>
          <span class="p-1 rounded bg-brand-50 text-brand-600">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
          </span>
        </div>
        <div class="text-xl font-extrabold text-brand-700 tracking-tight" id="kpi-primes-scg">0 FCFA</div>
        <div class="text-[11px] text-slate-500 mt-1 flex items-center justify-between">
          <span id="kpi-sub-primes-scg-label">Part du 100% :</span>
          <span class="font-bold text-brand-600" id="kpi-part-scg">0%</span>
        </div>
        <div class="absolute bottom-0 left-0 right-0 h-1 bg-brand-500"></div>
      </div>

      <!-- KPI 4: Commission -->
      <div class="bg-white rounded-xl border border-slate-200/90 p-4 shadow-xs relative overflow-hidden group hover:border-amber-300 transition-all">
        <div class="flex items-center justify-between text-slate-500 mb-1">
          <span class="text-xs font-semibold text-slate-600" id="kpi-label-commissions">Commission</span>
          <span class="p-1 rounded bg-amber-50 text-amber-600">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 9V7a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2m2 4h10a2 2 0 002-2v-6a2 2 0 00-2-2H9a2 2 0 00-2 2v6a2 2 0 002 2zm7-5a2 2 0 11-4 0 2 2 0 014 0z"/></svg>
          </span>
        </div>
        <div class="text-xl font-extrabold text-amber-800 tracking-tight" id="kpi-commissions">0 FCFA</div>
        <div class="text-[11px] text-slate-500 mt-1 flex items-center justify-between">
          <span>Taux moyen :</span>
          <span class="font-bold text-amber-700" id="kpi-taux-comm-moyen">0%</span>
        </div>
        <div class="absolute bottom-0 left-0 right-0 h-1 bg-amber-500"></div>
      </div>

      <!-- KPI 5: Prime Nette -->
      <div class="bg-white rounded-xl border border-slate-200/90 p-4 shadow-xs relative overflow-hidden group hover:border-emerald-300 transition-all">
        <div class="flex items-center justify-between text-slate-500 mb-1">
          <span class="text-xs font-semibold text-slate-600" id="kpi-label-primes-nettes">Prime Nette</span>
          <span class="p-1 rounded bg-emerald-50 text-emerald-600">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4M7.835 4.697a3.42 3.42 0 001.946-.806 3.42 3.42 0 014.438 0 3.42 3.42 0 001.946.806 3.42 3.42 0 013.138 3.138 3.42 3.42 0 00.806 1.946 3.42 3.42 0 010 4.438 3.42 3.42 0 00-.806 1.946 3.42 3.42 0 01-3.138 3.138 3.42 3.42 0 00-1.946.806 3.42 3.42 0 01-4.438 0 3.42 3.42 0 00-1.946-.806 3.42 3.42 0 01-3.138-3.138 3.42 3.42 0 00-.806-1.946 3.42 3.42 0 010-4.438 3.42 3.42 0 00.806-1.946 3.42 3.42 0 013.138-3.138z"/></svg>
          </span>
        </div>
        <div class="text-xl font-extrabold text-emerald-700 tracking-tight" id="kpi-primes-nettes">0 FCFA</div>
        <div class="text-[11px] text-slate-500 mt-1 flex items-center justify-between">
          <span>Net / Brut SCG :</span>
          <span class="font-bold text-emerald-600" id="kpi-ratio-net">0%</span>
        </div>
        <div class="absolute bottom-0 left-0 right-0 h-1 bg-emerald-500"></div>
      </div>

    </section>

    <!-- ========================================== -->
    <!-- TAB 1: VUE PRIME SOUSCRITE (VALEURS BRUTES) -->
    <!-- ========================================== -->
    <div id="tab-souscrite" class="tab-content space-y-6">
      
      <!-- Section Banner -->
      <div class="bg-gradient-to-r from-brand-900 to-brand-800 text-white rounded-xl p-4 shadow-sm flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 class="text-base font-bold flex items-center gap-2">
            <span>📊</span> Vue Prime souscrite — Valeurs Brutes des Primes SCG-Ré
          </h2>
          <p class="text-xs text-brand-200 mt-0.5">
            Analyse complète de l'ensemble de la production souscrite ARO, ventilée par Branche, Pays, Cédante et Courtier par ordre strictement décroissant.
          </p>
        </div>
        <span class="bg-brand-700 text-brand-100 text-xs px-3 py-1 rounded-full font-semibold border border-brand-500/40" id="badge-souscrite-total">
          Total : 0 FCFA
        </span>
      </div>

      <!-- 2x2 Grid : Branche & Pays -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Graphique 1: Prime souscrite par Branche -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-brand-500"></span>
                Prime souscrite par Branche
              </h3>
              <p class="text-[11px] text-slate-500">Classée par ordre décroissant (FCFA)</p>
            </div>
            <span class="text-xs font-semibold px-2 py-0.5 rounded bg-brand-50 text-brand-700 border border-brand-200" id="count-branches-souscrite">11 branches</span>
          </div>
          <div class="h-80 relative flex-1">
            <canvas id="chart-souscrite-branche"></canvas>
          </div>
        </div>

        <!-- Graphique 2: Prime souscrite par Pays -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span>
                Prime souscrite par Pays
              </h3>
              <p class="text-[11px] text-slate-500">Lieu du risque — Classé par ordre décroissant (FCFA)</p>
            </div>
            <span class="text-xs font-semibold px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200" id="count-pays-souscrite">0 pays</span>
          </div>
          <div class="h-80 relative flex-1">
            <canvas id="chart-souscrite-pays"></canvas>
          </div>
        </div>

      </div>

      <!-- 2x2 Grid : Cédante & Courtier (avec Bascule Graphique / Tableau Complet) -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Bloc 3: Prime souscrite par Cédante -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-3 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
                Prime souscrite par Cédante (Groupes standardisés)
              </h3>
              <p class="text-[11px] text-slate-500">Classée par ordre décroissant (FCFA)</p>
            </div>
            <!-- Toggle Graphique / Tableau -->
            <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
              <button id="btn-toggle-souscrite-cedante-chart" class="px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all">
                Graphique Top 10
              </button>
              <button id="btn-toggle-souscrite-cedante-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                Tableau Complet
              </button>
            </div>
          </div>

          <!-- Vue Graphique -->
          <div id="view-souscrite-cedante-chart" class="h-80 relative flex-1">
            <canvas id="chart-souscrite-cedante"></canvas>
          </div>

          <!-- Vue Tableau Complet -->
          <div id="view-souscrite-cedante-table" class="hidden flex-1 overflow-y-auto custom-scrollbar max-h-80">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-50 sticky top-0 text-[11px] uppercase tracking-wider text-slate-600 border-b border-slate-200">
                <tr>
                  <th class="py-2 px-2.5 font-bold w-12 text-center">Rang</th>
                  <th class="py-2 px-2.5 font-bold">Cédante (Groupe)</th>
                  <th class="py-2 px-2.5 font-bold text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-2.5 font-bold text-right w-20">Part %</th>
                </tr>
              </thead>
              <tbody id="table-souscrite-cedante-body" class="divide-y divide-slate-100">
                <!-- Rendu Dynamique -->
              </tbody>
            </table>
          </div>
        </div>

        <!-- Bloc 4: Prime souscrite par Courtier -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-3 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-indigo-500"></span>
                Prime souscrite par Courtier / Apporteur
              </h3>
              <p class="text-[11px] text-slate-500">Apporteurs standardisés — Ordre décroissant (FCFA)</p>
            </div>
            <!-- Toggle Graphique / Tableau -->
            <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
              <button id="btn-toggle-souscrite-courtier-chart" class="px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all">
                Graphique Top 10
              </button>
              <button id="btn-toggle-souscrite-courtier-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                Tableau Complet
              </button>
            </div>
          </div>

          <!-- Vue Graphique -->
          <div id="view-souscrite-courtier-chart" class="h-80 relative flex-1">
            <canvas id="chart-souscrite-courtier"></canvas>
          </div>

          <!-- Vue Tableau Complet -->
          <div id="view-souscrite-courtier-table" class="hidden flex-1 overflow-y-auto custom-scrollbar max-h-80">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-50 sticky top-0 text-[11px] uppercase tracking-wider text-slate-600 border-b border-slate-200">
                <tr>
                  <th class="py-2 px-2.5 font-bold w-12 text-center">Rang</th>
                  <th class="py-2 px-2.5 font-bold">Courtier / Apporteur</th>
                  <th class="py-2 px-2.5 font-bold text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-2.5 font-bold text-right w-20">Part %</th>
                </tr>
              </thead>
              <tbody id="table-souscrite-courtier-body" class="divide-y divide-slate-100">
                <!-- Rendu Dynamique -->
              </tbody>
            </table>
          </div>
        </div>

      </div>

    </div>

    <!-- ========================================== -->
    <!-- TAB 2: VUE PRIME SIGNÉE (BORDEREAUX REÇUS) -->
    <!-- ========================================== -->
    <div id="tab-signee" class="tab-content hidden space-y-6">
      
      <!-- Section Banner -->
      <div class="bg-gradient-to-r from-emerald-900 via-emerald-800 to-teal-900 text-white rounded-xl p-4 shadow-sm flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 class="text-base font-bold flex items-center gap-2">
            <span>✅</span> Vue Prime signée — Affaires avec Bordereaux Reçus
          </h2>
          <p class="text-xs text-emerald-200 mt-0.5">
            Analyse des affaires ARO fermes dont les bordereaux ont été effectivement reçus par la SCG-Ré, ventilées par ordre strictement décroissant.
          </p>
        </div>
        <div class="flex items-center gap-2">
          <span class="bg-emerald-700/80 text-emerald-100 text-xs px-3 py-1 rounded-full font-semibold border border-emerald-500/40" id="badge-signee-count">
            0 bordereaux reçus
          </span>
          <span class="bg-white/20 text-white text-xs px-3 py-1 rounded-full font-bold" id="badge-signee-total">
            Total : 0 FCFA
          </span>
        </div>
      </div>

      <!-- 2x2 Grid : Branche & Pays (Signée) -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Graphique 1: Prime signée par Branche -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span>
                Prime signée par Branche
              </h3>
              <p class="text-[11px] text-slate-500">Bordereaux reçus — Tri décroissant (FCFA)</p>
            </div>
            <span class="text-xs font-semibold px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200" id="count-branches-signee">11 branches</span>
          </div>
          <div class="h-80 relative flex-1">
            <canvas id="chart-signee-branche"></canvas>
          </div>
        </div>

        <!-- Graphique 2: Prime signée par Pays -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-teal-500"></span>
                Prime signée par Pays
              </h3>
              <p class="text-[11px] text-slate-500">Lieu du risque — Tri décroissant (FCFA)</p>
            </div>
            <span class="text-xs font-semibold px-2 py-0.5 rounded bg-teal-50 text-teal-700 border border-teal-200" id="count-pays-signee">0 pays</span>
          </div>
          <div class="h-80 relative flex-1">
            <canvas id="chart-signee-pays"></canvas>
          </div>
        </div>

      </div>

      <!-- 2x2 Grid : Cédante & Courtier (Signée avec Bascule Graphique / Tableau Complet) -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Bloc 3: Prime signée par Cédante -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-3 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-cyan-600"></span>
                Prime signée par Cédante (Groupes standardisés)
              </h3>
              <p class="text-[11px] text-slate-500">Bordereaux reçus — Tri décroissant (FCFA)</p>
            </div>
            <!-- Toggle Graphique / Tableau -->
            <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
              <button id="btn-toggle-signee-cedante-chart" class="px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all">
                Graphique Top 10
              </button>
              <button id="btn-toggle-signee-cedante-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                Tableau Complet
              </button>
            </div>
          </div>

          <!-- Vue Graphique -->
          <div id="view-signee-cedante-chart" class="h-80 relative flex-1">
            <canvas id="chart-signee-cedante"></canvas>
          </div>

          <!-- Vue Tableau Complet -->
          <div id="view-signee-cedante-table" class="hidden flex-1 overflow-y-auto custom-scrollbar max-h-80">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-50 sticky top-0 text-[11px] uppercase tracking-wider text-slate-600 border-b border-slate-200">
                <tr>
                  <th class="py-2 px-2.5 font-bold w-12 text-center">Rang</th>
                  <th class="py-2 px-2.5 font-bold">Cédante (Groupe)</th>
                  <th class="py-2 px-2.5 font-bold text-right">Prime Signée</th>
                  <th class="py-2 px-2.5 font-bold text-right w-20">Part %</th>
                </tr>
              </thead>
              <tbody id="table-signee-cedante-body" class="divide-y divide-slate-100">
                <!-- Rendu Dynamique -->
              </tbody>
            </table>
          </div>
        </div>

        <!-- Bloc 4: Prime signée par Courtier -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-3 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-emerald-600"></span>
                Prime signée par Courtier / Apporteur
              </h3>
              <p class="text-[11px] text-slate-500">Apporteurs standardisés — Tri décroissant (FCFA)</p>
            </div>
            <!-- Toggle Graphique / Tableau -->
            <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
              <button id="btn-toggle-signee-courtier-chart" class="px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all">
                Graphique Top 10
              </button>
              <button id="btn-toggle-signee-courtier-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                Tableau Complet
              </button>
            </div>
          </div>

          <!-- Vue Graphique -->
          <div id="view-signee-courtier-chart" class="h-80 relative flex-1">
            <canvas id="chart-signee-courtier"></canvas>
          </div>

          <!-- Vue Tableau Complet -->
          <div id="view-signee-courtier-table" class="hidden flex-1 overflow-y-auto custom-scrollbar max-h-80">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-50 sticky top-0 text-[11px] uppercase tracking-wider text-slate-600 border-b border-slate-200">
                <tr>
                  <th class="py-2 px-2.5 font-bold w-12 text-center">Rang</th>
                  <th class="py-2 px-2.5 font-bold">Courtier / Apporteur</th>
                  <th class="py-2 px-2.5 font-bold text-right">Prime Signée</th>
                  <th class="py-2 px-2.5 font-bold text-right w-20">Part %</th>
                </tr>
              </thead>
              <tbody id="table-signee-courtier-body" class="divide-y divide-slate-100">
                <!-- Rendu Dynamique -->
              </tbody>
            </table>
          </div>
        </div>

      </div>

    </div>

    <!-- ========================================== -->
    <!-- TAB 3: EXPLORATEUR DE DONNÉES & EXPORT -->
    <!-- ========================================== -->
    <div id="tab-data" class="tab-content hidden space-y-4">
      
      <!-- Toolbar -->
      <div class="bg-white rounded-xl border border-slate-200 p-4 shadow-xs flex flex-wrap items-center justify-between gap-3">
        <div class="flex items-center gap-3 flex-1 min-w-[280px]">
          <div class="relative w-full max-w-md">
            <input type="text" id="table-search" placeholder="Rechercher une affaire, cédante, pays, courtier..." class="w-full pl-9 pr-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-xs focus:ring-2 focus:ring-brand-500 focus:border-brand-500">
            <svg class="w-4 h-4 text-slate-400 absolute left-2.5 top-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>
          </div>
          <span class="text-xs text-slate-500 whitespace-nowrap" id="table-records-found">0 polices</span>
        </div>

        <div class="flex items-center gap-2">
          <label class="text-xs text-slate-500">Lignes :</label>
          <select id="table-page-size" class="bg-slate-50 border border-slate-300 text-xs rounded-lg px-2 py-1">
            <option value="25">25</option>
            <option value="50">50</option>
            <option value="100">100</option>
            <option value="ALL">Tout</option>
          </select>
          <button id="btn-export-table" class="inline-flex items-center gap-1.5 px-3 py-1.5 bg-brand-600 hover:bg-brand-500 text-white rounded-lg text-xs font-semibold shadow-xs transition-all">
            <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>
            Exporter ce tableau (.xlsx)
          </button>
        </div>
      </div>

      <!-- Data Table Card -->
      <div class="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div class="overflow-x-auto custom-scrollbar">
          <table class="w-full text-left text-xs border-collapse">
            <thead class="bg-slate-50 text-[11px] uppercase tracking-wider text-slate-600 border-b border-slate-200 select-none">
              <tr>
                <th class="py-3 px-3 font-bold cursor-pointer hover:bg-slate-100" data-col="id">N° ↕</th>
                <th class="py-3 px-3 font-bold cursor-pointer hover:bg-slate-100" data-col="dateSouscription">Date Souscr. ↕</th>
                <th class="py-3 px-3 font-bold cursor-pointer hover:bg-slate-100" data-col="affaire">Affaire / Police ↕</th>
                <th class="py-3 px-3 font-bold cursor-pointer hover:bg-slate-100" data-col="cedante">Cédante (Groupe) ↕</th>
                <th class="py-3 px-3 font-bold cursor-pointer hover:bg-slate-100" data-col="apporteur">Courtier ↕</th>
                <th class="py-3 px-3 font-bold cursor-pointer hover:bg-slate-100" data-col="pays">Pays ↕</th>
                <th class="py-3 px-3 font-bold cursor-pointer hover:bg-slate-100" data-col="branche">Branche ↕</th>
                <th class="py-3 px-3 font-bold text-right cursor-pointer hover:bg-slate-100" data-col="capitauxSCG">Capitaux SCG ↕</th>
                <th class="py-3 px-3 font-bold text-right cursor-pointer hover:bg-slate-100" data-col="primesSCG">Prime SCG ↕</th>
                <th class="py-3 px-3 font-bold text-right cursor-pointer hover:bg-slate-100" data-col="commission">Commission ↕</th>
                <th class="py-3 px-3 font-bold text-right cursor-pointer hover:bg-slate-100" data-col="primesNettes">Prime Nette ↕</th>
                <th class="py-3 px-3 font-bold text-center cursor-pointer hover:bg-slate-100" data-col="etatBordereau">Bordereau ↕</th>
              </tr>
            </thead>
            <tbody id="data-table-body" class="divide-y divide-slate-100 text-slate-700 font-normal">
              <!-- Dynamically populated -->
            </tbody>
          </table>
        </div>

        <!-- Pagination -->
        <div class="px-4 py-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between text-xs text-slate-600 gap-2">
          <div id="table-pagination-info">Page 1 sur 1</div>
          <div class="flex items-center gap-1.5" id="table-pagination-buttons">
            <!-- Buttons injected -->
          </div>
        </div>
      </div>

    </div>

  </main>

  <!-- FOOTER (PLEINE LARGEUR) -->
  <footer class="bg-white border-t border-slate-200 py-3.5 no-print mt-auto">
    <div class="w-full px-4 sm:px-6 lg:px-8 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-500">
      <div class="flex items-center gap-2">
        <span class="w-2 h-2 rounded-full bg-emerald-500"></span>
        <span class="font-bold text-slate-700">SCG-Ré</span> — Société Commerciale Gabonaise de Réassurance
      </div>
      <div>
        <span>Tableau de Bord ARO — Cessions FAC — Pleine Largeur</span>
      </div>
    </div>
  </footer>

  <!-- MODAL: UPLOAD EXCEL FILE -->
  <div id="upload-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-4 hidden no-print">
    <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full p-6 border border-slate-200">
      
      <div class="flex items-center justify-between pb-3 border-b border-slate-100 mb-4">
        <div class="flex items-center gap-2">
          <div class="p-2 rounded-lg bg-brand-50 text-brand-600">
            <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>
          </div>
          <div>
            <h3 class="text-sm font-bold text-slate-900">Charger la production ARO actualisée</h3>
            <p class="text-[11px] text-slate-500">Classeur Excel brut (.xlsx, .xls) — Traitement 100% automatique</p>
          </div>
        </div>
        <button id="btn-close-upload" class="p-1 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
        </button>
      </div>

      <!-- Drag & Drop Zone -->
      <div id="drop-zone" class="border-2 border-dashed border-slate-300 hover:border-brand-500 hover:bg-brand-50/30 rounded-xl p-8 text-center transition-all cursor-pointer">
        <input type="file" id="file-input" accept=".xlsx, .xls, .xlsm" class="hidden">
        <svg class="w-10 h-10 text-slate-400 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
        <p class="text-xs font-semibold text-slate-700">Cliquez pour parcourir ou glissez votre classeur ARO ici</p>
        <p class="text-[11px] text-slate-400 mt-1">Formats acceptés : Classeur Excel (.xlsx, .xls, .xlsm)</p>
      </div>

      <div class="mt-4 p-3 bg-emerald-50 rounded-xl border border-emerald-200 text-[11px] text-emerald-800 flex items-start gap-2">
        <svg class="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
        <span><strong>Traitements 100% automatiques :</strong> standardisation des 11 branches, normalisation des pays, correction des apporteurs et regroupement des cédantes par Groupe d'assurance.</span>
      </div>

    </div>
  </div>

  <!-- TOAST NOTIFICATION -->
  <div id="toast" class="fixed bottom-5 right-5 bg-slate-900 text-white text-xs font-semibold px-4 py-3 rounded-xl shadow-2xl flex items-center gap-2 transform translate-y-20 opacity-0 transition-all duration-300 z-50">
    <svg class="w-4 h-4 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"/></svg>
    <span id="toast-message">Action effectuée</span>
  </div>

  <!-- EMBEDDED DATASET & CORE APPLICATION LOGIC -->
  <script>
    // Register ChartDataLabels plugin globally
    Chart.register(ChartDataLabels);

    // INITIAL EMBEDDED CLEANED DATA
    const INITIAL_RECORDS = __INITIAL_RECORDS__;
    const INITIAL_CLEANING_LOG = __INITIAL_CLEANING_LOG__;

    // ACTIVE STATE
    let allRecords = [...INITIAL_RECORDS];
    let cleaningLog = [...INITIAL_CLEANING_LOG];
    let filteredRecords = [...allRecords];
    let currentTabId = 'tab-souscrite';

    // CHART INSTANCES
    let charts = {};

    // TABLE STATE (Tab Explorateur)
    let currentPage = 1;
    let pageSize = 25;
    let sortColumn = 'id';
    let sortDirection = 1; // 1 = asc, -1 = desc
    let searchTerm = '';

    // ==========================================
    // ROBUST FORMATTERS (SÉPARATEURS D'ESPACES)
    // ==========================================
    const formatFCFA = (val) => {
      if (val === null || val === undefined || isNaN(val)) return '0 FCFA';
      const rounded = Math.round(val);
      return rounded.toString().replace(/\\B(?=(\\d{3})+(?!\\d))/g, ' ') + ' FCFA';
    };

    const formatNum = (val) => {
      if (val === null || val === undefined || isNaN(val)) return '0';
      const rounded = Math.round(val);
      return rounded.toString().replace(/\\B(?=(\\d{3})+(?!\\d))/g, ' ');
    };

    const formatPercent = (val) => {
      if (val === null || val === undefined || isNaN(val)) return '0%';
      return (val * 100).toFixed(1) + '%';
    };

    // Compact format for bar charts datalabels (e.g. 1 853 M or 45 k)
    const formatBarValue = (val) => {
      if (!val || isNaN(val)) return '0';
      const abs = Math.abs(val);
      if (abs >= 1e9) {
        return (val / 1e9).toFixed(2).replace('.', ',') + ' Mrd';
      } else if (abs >= 1e6) {
        return (val / 1e6).toFixed(1).replace('.', ',') + ' M';
      } else if (abs >= 1e3) {
        return (val / 1e3).toFixed(0).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ' ') + ' k';
      }
      return formatNum(val);
    };

    // TOAST NOTIFIER
    function showToast(msg) {
      const toast = document.getElementById('toast');
      const toastMsg = document.getElementById('toast-message');
      toastMsg.textContent = msg;
      toast.classList.remove('translate-y-20', 'opacity-0');
      setTimeout(() => {
        toast.classList.add('translate-y-20', 'opacity-0');
      }, 3500);
    }

    // ==========================================
    // INITIALIZATION ON DOM READY
    // ==========================================
    document.addEventListener('DOMContentLoaded', () => {
      setupTabs();
      setupFilters();
      setupUploadModal();
      setupTableControls();
      setupViewToggles();
      renderAll();
    });

    // ==========================================
    // TAB SWITCHING & DYNAMIC KPI CONTEXT
    // ==========================================
    function setupTabs() {
      const tabBtns = document.querySelectorAll('.tab-btn');
      tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
          const targetTab = btn.getAttribute('data-tab');
          currentTabId = targetTab;
          
          tabBtns.forEach(b => {
            b.classList.remove('active', 'bg-brand-50', 'text-brand-700', 'border-brand-200/80');
            b.classList.add('text-slate-600');
          });
          btn.classList.add('active', 'bg-brand-50', 'text-brand-700', 'border-brand-200/80');
          btn.classList.remove('text-slate-600');

          document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
          const activeContent = document.getElementById(targetTab);
          if (activeContent) activeContent.classList.remove('hidden');

          // Re-render KPIs according to current active tab
          renderKPIs();

          // Resize active charts
          setTimeout(() => {
            Object.values(charts).forEach(c => {
              if (c && typeof c.resize === 'function') c.resize();
            });
          }, 50);
        });
      });
    }

    // ==========================================
    // TOGGLES GRAPHIQUE / TABLEAU POUR CÉDANTES & COURTIERS
    // ==========================================
    function setupViewToggles() {
      // Souscrite Cédante
      const btnSCChart = document.getElementById('btn-toggle-souscrite-cedante-chart');
      const btnSCTable = document.getElementById('btn-toggle-souscrite-cedante-table');
      const viewSCChart = document.getElementById('view-souscrite-cedante-chart');
      const viewSCTable = document.getElementById('view-souscrite-cedante-table');
      
      btnSCChart.addEventListener('click', () => {
        btnSCChart.className = 'px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all';
        btnSCTable.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        viewSCChart.classList.remove('hidden');
        viewSCTable.classList.add('hidden');
        if (charts['chart-souscrite-cedante']) charts['chart-souscrite-cedante'].resize();
      });
      btnSCTable.addEventListener('click', () => {
        btnSCTable.className = 'px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all';
        btnSCChart.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        viewSCTable.classList.remove('hidden');
        viewSCChart.classList.add('hidden');
      });

      // Souscrite Courtier
      const btnSKChart = document.getElementById('btn-toggle-souscrite-courtier-chart');
      const btnSKTable = document.getElementById('btn-toggle-souscrite-courtier-table');
      const viewSKChart = document.getElementById('view-souscrite-courtier-chart');
      const viewSKTable = document.getElementById('view-souscrite-courtier-table');

      btnSKChart.addEventListener('click', () => {
        btnSKChart.className = 'px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all';
        btnSKTable.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        viewSKChart.classList.remove('hidden');
        viewSKTable.classList.add('hidden');
        if (charts['chart-souscrite-courtier']) charts['chart-souscrite-courtier'].resize();
      });
      btnSKTable.addEventListener('click', () => {
        btnSKTable.className = 'px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all';
        btnSKChart.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        viewSKTable.classList.remove('hidden');
        viewSKChart.classList.add('hidden');
      });

      // Signée Cédante
      const btnSgCChart = document.getElementById('btn-toggle-signee-cedante-chart');
      const btnSgCTable = document.getElementById('btn-toggle-signee-cedante-table');
      const viewSgCChart = document.getElementById('view-signee-cedante-chart');
      const viewSgCTable = document.getElementById('view-signee-cedante-table');

      btnSgCChart.addEventListener('click', () => {
        btnSgCChart.className = 'px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all';
        btnSgCTable.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        viewSgCChart.classList.remove('hidden');
        viewSgCTable.classList.add('hidden');
        if (charts['chart-signee-cedante']) charts['chart-signee-cedante'].resize();
      });
      btnSgCTable.addEventListener('click', () => {
        btnSgCTable.className = 'px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all';
        btnSgCChart.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        viewSgCTable.classList.remove('hidden');
        viewSgCChart.classList.add('hidden');
      });

      // Signée Courtier
      const btnSgKChart = document.getElementById('btn-toggle-signee-courtier-chart');
      const btnSgKTable = document.getElementById('btn-toggle-signee-courtier-table');
      const viewSgKChart = document.getElementById('view-signee-courtier-chart');
      const viewSgKTable = document.getElementById('view-signee-courtier-table');

      btnSgKChart.addEventListener('click', () => {
        btnSgKChart.className = 'px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all';
        btnSgKTable.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        viewSgKChart.classList.remove('hidden');
        viewSgKTable.classList.add('hidden');
        if (charts['chart-signee-courtier']) charts['chart-signee-courtier'].resize();
      });
      btnSgKTable.addEventListener('click', () => {
        btnSgKTable.className = 'px-2.5 py-1 rounded-md font-semibold bg-white text-slate-800 shadow-xs transition-all';
        btnSgKChart.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        viewSgKTable.classList.remove('hidden');
        viewSgKChart.classList.add('hidden');
      });
    }

    // ==========================================
    // FILTERS (FILTRE UNIQUE BRANCHE)
    // ==========================================
    function setupFilters() {
      populateFilterOptions();

      document.getElementById('filter-branche').addEventListener('change', applyFilters);

      document.getElementById('btn-reset-filters').addEventListener('click', () => {
        document.getElementById('filter-branche').value = 'ALL';
        applyFilters();
        showToast('Filtre réinitialisé');
      });
    }

    function populateFilterOptions() {
      const brancheSet = new Set();
      allRecords.forEach(r => {
        if (r.branche) brancheSet.add(r.branche);
      });

      const selBranche = document.getElementById('filter-branche');
      selBranche.innerHTML = '<option value="ALL">Toutes les branches (' + brancheSet.size + ')</option>';
      Array.from(brancheSet).sort().forEach(b => {
        const count = allRecords.filter(r => r.branche === b).length;
        selBranche.innerHTML += `<option value="${b}">${b} (${count})</option>`;
      });

      document.getElementById('total-count-badge').textContent = allRecords.length;
    }

    function applyFilters() {
      const selB = document.getElementById('filter-branche').value;

      filteredRecords = allRecords.filter(r => {
        if (selB !== 'ALL' && r.branche !== selB) return false;
        return true;
      });

      document.getElementById('filtered-count').textContent = filteredRecords.length;
      currentPage = 1;
      renderAll();
    }

    // ==========================================
    // MAIN RENDER CALL
    // ==========================================
    function renderAll() {
      renderKPIs();
      renderVueSouscrite();
      renderVueSignee();
      renderDataTable();
    }

    // ==========================================
    // 1. RENDER 5 PERMANENT KPIS (DYNAMIQUES PAR ONGLET)
    // ==========================================
    function renderKPIs() {
      const isSigneeMode = (currentTabId === 'tab-signee');
      const activeSubset = isSigneeMode 
        ? filteredRecords.filter(r => r.etatBordereau === 'REÇU')
        : filteredRecords;

      const totalCapitaux100 = activeSubset.reduce((acc, r) => acc + (r.capitaux100 || 0), 0);
      const totalCapitauxSCG = activeSubset.reduce((acc, r) => acc + (r.capitauxSCG || 0), 0);
      const totalPrimes100 = activeSubset.reduce((acc, r) => acc + (r.primes100 || 0), 0);
      const totalPrimesSCG = activeSubset.reduce((acc, r) => acc + (r.primesSCG || 0), 0);
      const totalCommissions = activeSubset.reduce((acc, r) => acc + (r.commission || 0), 0);
      const totalPrimesNettes = activeSubset.reduce((acc, r) => acc + (r.primesNettes || 0), 0);

      const partCapitauxSCG = totalCapitaux100 > 0 ? (totalCapitauxSCG / totalCapitaux100) : 0;
      const partPrimesSCG = totalPrimes100 > 0 ? (totalPrimesSCG / totalPrimes100) : 0;
      const avgCommRate = totalPrimesSCG > 0 ? (totalCommissions / totalPrimesSCG) : 0;
      const ratioNet = totalPrimesSCG > 0 ? (totalPrimesNettes / totalPrimesSCG) : 0;

      // Update Card 1: Capitaux Totaux 100%
      document.getElementById('kpi-label-capitaux-100').textContent = isSigneeMode 
        ? 'Capitaux Totaux 100% (Signés)' 
        : 'Capitaux Totaux (100%)';
      document.getElementById('kpi-capitaux-100').textContent = formatFCFA(totalCapitaux100);
      document.getElementById('kpi-sub-capitaux-100').textContent = isSigneeMode 
        ? activeSubset.length + ' bordereaux reçus'
        : '100% Marché';

      // Update Card 2: Engagement SCG-Ré
      document.getElementById('kpi-label-capitaux-scg').textContent = isSigneeMode 
        ? 'Engagement SCG-Ré (Signé)' 
        : 'Engagement SCG-Ré';
      document.getElementById('kpi-capitaux-scg').textContent = formatFCFA(totalCapitauxSCG);
      document.getElementById('kpi-part-capitaux-scg').textContent = formatPercent(partCapitauxSCG);

      // Update Card 3: Prime SCG-Ré
      document.getElementById('kpi-label-primes-scg').textContent = isSigneeMode 
        ? 'Prime SCG-Ré (Signée)' 
        : 'Prime SCG-Ré';
      document.getElementById('kpi-primes-scg').textContent = formatFCFA(totalPrimesSCG);
      document.getElementById('kpi-part-scg').textContent = formatPercent(partPrimesSCG);

      // Update Card 4: Commission
      document.getElementById('kpi-label-commissions').textContent = isSigneeMode 
        ? 'Commission (Signée)' 
        : 'Commission';
      document.getElementById('kpi-commissions').textContent = formatFCFA(totalCommissions);
      document.getElementById('kpi-taux-comm-moyen').textContent = formatPercent(avgCommRate);

      // Update Card 5: Prime Nette
      document.getElementById('kpi-label-primes-nettes').textContent = isSigneeMode 
        ? 'Prime Nette (Signée)' 
        : 'Prime Nette';
      document.getElementById('kpi-primes-nettes').textContent = formatFCFA(totalPrimesNettes);
      document.getElementById('kpi-ratio-net').textContent = formatPercent(ratioNet);

      document.getElementById('header-records-count').textContent = allRecords.length + ' affaires chargées';
    }

    // ==========================================
    // 2. VUE PRIME SOUSCRITE (VALEURS BRUTES)
    // ==========================================
    function renderVueSouscrite() {
      const records = filteredRecords;
      const totalSouscrite = records.reduce((acc, r) => acc + (r.primesSCG || 0), 0);
      document.getElementById('badge-souscrite-total').textContent = 'Total : ' + formatFCFA(totalSouscrite);

      // --- A. Prime souscrite par Branche (Ordre décroissant) ---
      const brancheMap = {};
      records.forEach(r => {
        const b = r.branche || 'AUTRES';
        brancheMap[b] = (brancheMap[b] || 0) + (r.primesSCG || 0);
      });
      const sortedBranches = Object.entries(brancheMap).sort((a, b) => b[1] - a[1]);
      document.getElementById('count-branches-souscrite').textContent = sortedBranches.length + ' branches';

      renderHorizontalBarChart('chart-souscrite-branche', {
        labels: sortedBranches.map(x => x[0]),
        values: sortedBranches.map(x => x[1]),
        color: '#0284c7',
        datasetLabel: 'Prime souscrite'
      });

      // --- B. Prime souscrite par Pays (Ordre décroissant) ---
      const paysMap = {};
      records.forEach(r => {
        const p = r.pays || 'NON DÉFINI';
        paysMap[p] = (paysMap[p] || 0) + (r.primesSCG || 0);
      });
      const sortedPays = Object.entries(paysMap).sort((a, b) => b[1] - a[1]);
      document.getElementById('count-pays-souscrite').textContent = sortedPays.length + ' pays';

      renderHorizontalBarChart('chart-souscrite-pays', {
        labels: sortedPays.map(x => x[0]),
        values: sortedPays.map(x => x[1]),
        color: '#10b981',
        datasetLabel: 'Prime souscrite'
      });

      // --- C. Prime souscrite par Cédante (Groupes standardisés) ---
      const cedanteMap = {};
      records.forEach(r => {
        const c = r.cedante || 'NON DÉFINI';
        cedanteMap[c] = (cedanteMap[c] || 0) + (r.primesSCG || 0);
      });
      const sortedCedantes = Object.entries(cedanteMap).sort((a, b) => b[1] - a[1]);

      // Top 10 pour le graphique
      const top10Cedantes = sortedCedantes.slice(0, 10);
      renderHorizontalBarChart('chart-souscrite-cedante', {
        labels: top10Cedantes.map(x => x[0]),
        values: top10Cedantes.map(x => x[1]),
        color: '#d97706',
        datasetLabel: 'Prime souscrite'
      });

      // Tableau complet pour la bascule
      renderRankedTable('table-souscrite-cedante-body', sortedCedantes, totalSouscrite);

      // --- D. Prime souscrite par Courtier / Apporteur ---
      const courtierMap = {};
      records.forEach(r => {
        const a = r.apporteur || 'NON SPÉCIFIÉ';
        courtierMap[a] = (courtierMap[a] || 0) + (r.primesSCG || 0);
      });
      const sortedCourtiers = Object.entries(courtierMap).sort((a, b) => b[1] - a[1]);

      // Top 10 pour le graphique
      const top10Courtiers = sortedCourtiers.slice(0, 10);
      renderHorizontalBarChart('chart-souscrite-courtier', {
        labels: top10Courtiers.map(x => x[0]),
        values: top10Courtiers.map(x => x[1]),
        color: '#6366f1',
        datasetLabel: 'Prime souscrite'
      });

      // Tableau complet pour la bascule
      renderRankedTable('table-souscrite-courtier-body', sortedCourtiers, totalSouscrite);
    }

    // ==========================================
    // 3. VUE PRIME SIGNÉE (BORDEREAUX REÇUS)
    // ==========================================
    function renderVueSignee() {
      const records = filteredRecords.filter(r => r.etatBordereau === 'REÇU');
      const totalSignee = records.reduce((acc, r) => acc + (r.primesSCG || 0), 0);

      document.getElementById('badge-signee-count').textContent = records.length + ' bordereaux reçus';
      document.getElementById('badge-signee-total').textContent = 'Total : ' + formatFCFA(totalSignee);

      // --- A. Prime signée par Branche (Ordre décroissant) ---
      const brancheMap = {};
      records.forEach(r => {
        const b = r.branche || 'AUTRES';
        brancheMap[b] = (brancheMap[b] || 0) + (r.primesSCG || 0);
      });
      const sortedBranches = Object.entries(brancheMap).sort((a, b) => b[1] - a[1]);
      document.getElementById('count-branches-signee').textContent = sortedBranches.length + ' branches';

      renderHorizontalBarChart('chart-signee-branche', {
        labels: sortedBranches.map(x => x[0]),
        values: sortedBranches.map(x => x[1]),
        color: '#059669',
        datasetLabel: 'Prime signée'
      });

      // --- B. Prime signée par Pays (Ordre décroissant) ---
      const paysMap = {};
      records.forEach(r => {
        const p = r.pays || 'NON DÉFINI';
        paysMap[p] = (paysMap[p] || 0) + (r.primesSCG || 0);
      });
      const sortedPays = Object.entries(paysMap).sort((a, b) => b[1] - a[1]);
      document.getElementById('count-pays-signee').textContent = sortedPays.length + ' pays';

      renderHorizontalBarChart('chart-signee-pays', {
        labels: sortedPays.map(x => x[0]),
        values: sortedPays.map(x => x[1]),
        color: '#0d9488',
        datasetLabel: 'Prime signée'
      });

      // --- C. Prime signée par Cédante (Groupes standardisés) ---
      const cedanteMap = {};
      records.forEach(r => {
        const c = r.cedante || 'NON DÉFINI';
        cedanteMap[c] = (cedanteMap[c] || 0) + (r.primesSCG || 0);
      });
      const sortedCedantes = Object.entries(cedanteMap).sort((a, b) => b[1] - a[1]);

      // Top 10 pour le graphique
      const top10Cedantes = sortedCedantes.slice(0, 10);
      renderHorizontalBarChart('chart-signee-cedante', {
        labels: top10Cedantes.map(x => x[0]),
        values: top10Cedantes.map(x => x[1]),
        color: '#0891b2',
        datasetLabel: 'Prime signée'
      });

      // Tableau complet pour la bascule
      renderRankedTable('table-signee-cedante-body', sortedCedantes, totalSignee);

      // --- D. Prime signée par Courtier / Apporteur ---
      const courtierMap = {};
      records.forEach(r => {
        const a = r.apporteur || 'NON SPÉCIFIÉ';
        courtierMap[a] = (courtierMap[a] || 0) + (r.primesSCG || 0);
      });
      const sortedCourtiers = Object.entries(courtierMap).sort((a, b) => b[1] - a[1]);

      // Top 10 pour le graphique
      const top10Courtiers = sortedCourtiers.slice(0, 10);
      renderHorizontalBarChart('chart-signee-courtier', {
        labels: top10Courtiers.map(x => x[0]),
        values: top10Courtiers.map(x => x[1]),
        color: '#10b981',
        datasetLabel: 'Prime signée'
      });

      // Tableau complet pour la bascule
      renderRankedTable('table-signee-courtier-body', sortedCourtiers, totalSignee);
    }

    // ==========================================
    // REUSABLE HORIZONTAL BAR CHART BUILDER
    // Avec montants au-dessus / au bout des barres
    // ==========================================
    function renderHorizontalBarChart(canvasId, config) {
      const { labels, values, color, datasetLabel } = config;
      const ctx = document.getElementById(canvasId);
      if (!ctx) return;

      if (charts[canvasId]) {
        charts[canvasId].destroy();
      }

      const maxValue = Math.max(...values, 0);

      charts[canvasId] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [{
            label: datasetLabel,
            data: values,
            backgroundColor: color,
            borderRadius: 5,
            borderSkipped: false,
          }]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          layout: {
            padding: {
              right: 65, // Espace réservé pour l'étiquette au bout de la barre
              left: 5,
              top: 5,
              bottom: 5
            }
          },
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: (ctx) => `${ctx.dataset.label} : ${formatFCFA(ctx.raw)}`
              }
            },
            datalabels: {
              anchor: 'end',
              align: 'right',
              offset: 4,
              color: '#334155',
              font: {
                size: 10,
                weight: '700',
                family: 'Plus Jakarta Sans'
              },
              formatter: (val) => {
                if (!val || val === 0) return '';
                return formatBarValue(val);
              }
            }
          },
          scales: {
            x: {
              suggestedMax: maxValue > 0 ? maxValue * 1.15 : 100,
              grid: { color: '#f1f5f9' },
              ticks: {
                font: { size: 10 },
                color: '#64748b',
                callback: (v) => formatBarValue(v)
              }
            },
            y: {
              grid: { display: false },
              ticks: {
                font: { size: 10, weight: '600' },
                color: '#334155',
                autoSkip: false
              }
            }
          }
        }
      });
    }

    // Helper pour remplir les tableaux complets classés (Cédantes & Courtiers)
    function renderRankedTable(tbodyId, sortedList, total) {
      const tbody = document.getElementById(tbodyId);
      if (!tbody) return;

      if (sortedList.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" class="py-4 text-center text-slate-400">Aucune donnée disponible</td></tr>';
        return;
      }

      tbody.innerHTML = sortedList.map((item, idx) => {
        const part = total > 0 ? (item[1] / total) : 0;
        const partWidth = Math.min(100, Math.max(2, Math.round(part * 100)));
        return `
          <tr class="hover:bg-slate-50 transition-colors">
            <td class="py-2 px-2.5 font-bold text-center text-slate-500">${idx + 1}</td>
            <td class="py-2 px-2.5 font-semibold text-slate-800">
              <div class="truncate max-w-[220px]" title="${item[0]}">${item[0]}</div>
              <div class="w-full bg-slate-100 rounded-full h-1 mt-1">
                <div class="bg-brand-500 h-1 rounded-full" style="width: ${partWidth}%"></div>
              </div>
            </td>
            <td class="py-2 px-2.5 font-extrabold text-right text-slate-900 whitespace-nowrap">${formatFCFA(item[1])}</td>
            <td class="py-2 px-2.5 font-semibold text-right text-slate-600 whitespace-nowrap">${formatPercent(part)}</td>
          </tr>
        `;
      }).join('');
    }

    // ==========================================
    // 4. DATA EXPLORER & EXPORT (ONGLET 3)
    // ==========================================
    function setupTableControls() {
      document.getElementById('table-search').addEventListener('input', (e) => {
        searchTerm = e.target.value.toLowerCase().trim();
        currentPage = 1;
        renderDataTable();
      });

      document.getElementById('table-page-size').addEventListener('change', (e) => {
        pageSize = e.target.value === 'ALL' ? 999999 : parseInt(e.target.value);
        currentPage = 1;
        renderDataTable();
      });

      document.querySelectorAll('th[data-col]').forEach(th => {
        th.addEventListener('click', () => {
          const col = th.getAttribute('data-col');
          if (sortColumn === col) {
            sortDirection = -sortDirection;
          } else {
            sortColumn = col;
            sortDirection = 1;
          }
          renderDataTable();
        });
      });

      document.getElementById('btn-export-table').addEventListener('click', exportCurrentTable);
      document.getElementById('btn-export-cleaned').addEventListener('click', exportCurrentTable);
    }

    function renderDataTable() {
      let data = [...filteredRecords];

      // Text search
      if (searchTerm) {
        data = data.filter(r => {
          return (r.affaire && r.affaire.toLowerCase().includes(searchTerm)) ||
                 (r.cedante && r.cedante.toLowerCase().includes(searchTerm)) ||
                 (r.apporteur && r.apporteur.toLowerCase().includes(searchTerm)) ||
                 (r.pays && r.pays.toLowerCase().includes(searchTerm)) ||
                 (r.branche && r.branche.toLowerCase().includes(searchTerm)) ||
                 (r.etatBordereau && r.etatBordereau.toLowerCase().includes(searchTerm)) ||
                 (r.id && r.id.toString().includes(searchTerm));
        });
      }

      document.getElementById('table-records-found').textContent = data.length + ' polices trouvées';

      // Sort
      data.sort((a, b) => {
        let vA = a[sortColumn];
        let vB = b[sortColumn];
        if (vA === null || vA === undefined) vA = '';
        if (vB === null || vB === undefined) vB = '';
        if (typeof vA === 'number' && typeof vB === 'number') {
          return (vA - vB) * sortDirection;
        }
        return vA.toString().localeCompare(vB.toString()) * sortDirection;
      });

      // Pagination
      const totalPages = Math.ceil(data.length / pageSize) || 1;
      if (currentPage > totalPages) currentPage = totalPages;
      const startIdx = (currentPage - 1) * pageSize;
      const pagedData = data.slice(startIdx, startIdx + pageSize);

      const tbody = document.getElementById('data-table-body');
      if (pagedData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="12" class="py-6 text-center text-slate-400">Aucune police ne correspond aux filtres</td></tr>';
      } else {
        tbody.innerHTML = pagedData.map(r => {
          const isRecu = (r.etatBordereau === 'REÇU');
          return `
            <tr class="hover:bg-slate-50 transition-colors">
              <td class="py-2.5 px-3 font-semibold text-slate-500">${r.id}</td>
              <td class="py-2.5 px-3 text-slate-600 whitespace-nowrap">${r.dateSouscription || '-'}</td>
              <td class="py-2.5 px-3 font-bold text-slate-900 max-w-xs truncate" title="${r.affaire || ''}">${r.affaire || '-'}</td>
              <td class="py-2.5 px-3 font-medium text-slate-700 max-w-[180px] truncate" title="${r.cedante || ''}">${r.cedante || '-'}</td>
              <td class="py-2.5 px-3 text-slate-600 max-w-[160px] truncate" title="${r.apporteur || ''}">${r.apporteur || '-'}</td>
              <td class="py-2.5 px-3 text-slate-600 whitespace-nowrap">${r.pays || '-'}</td>
              <td class="py-2.5 px-3 text-slate-600 whitespace-nowrap">${r.branche || '-'}</td>
              <td class="py-2.5 px-3 font-medium text-right text-indigo-700 whitespace-nowrap">${formatFCFA(r.capitauxSCG)}</td>
              <td class="py-2.5 px-3 font-bold text-right text-brand-700 whitespace-nowrap">${formatFCFA(r.primesSCG)}</td>
              <td class="py-2.5 px-3 font-medium text-right text-amber-700 whitespace-nowrap">${formatFCFA(r.commission)}</td>
              <td class="py-2.5 px-3 font-bold text-right text-emerald-700 whitespace-nowrap">${formatFCFA(r.primesNettes)}</td>
              <td class="py-2.5 px-3 text-center whitespace-nowrap">
                <span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold ${isRecu ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}">
                  ${r.etatBordereau || 'EN ATTENTE'}
                </span>
              </td>
            </tr>
          `;
        }).join('');
      }

      // Pagination footer
      document.getElementById('table-pagination-info').textContent = 
        `Affichage ${Math.min(data.length, startIdx + 1)} - ${Math.min(data.length, startIdx + pageSize)} sur ${data.length} polices`;

      const btnContainer = document.getElementById('table-pagination-buttons');
      btnContainer.innerHTML = `
        <button onclick="changePage(1)" ${currentPage === 1 ? 'disabled class="opacity-40 cursor-not-allowed"' : ''} class="px-2 py-1 rounded border border-slate-300 bg-white hover:bg-slate-50 text-xs">«</button>
        <button onclick="changePage(${currentPage - 1})" ${currentPage === 1 ? 'disabled class="opacity-40 cursor-not-allowed"' : ''} class="px-2.5 py-1 rounded border border-slate-300 bg-white hover:bg-slate-50 text-xs">Précédent</button>
        <span class="px-2 py-1 font-bold text-xs">Page ${currentPage} / ${totalPages}</span>
        <button onclick="changePage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled class="opacity-40 cursor-not-allowed"' : ''} class="px-2.5 py-1 rounded border border-slate-300 bg-white hover:bg-slate-50 text-xs">Suivant</button>
        <button onclick="changePage(${totalPages})" ${currentPage === totalPages ? 'disabled class="opacity-40 cursor-not-allowed"' : ''} class="px-2.5 py-1 rounded border border-slate-300 bg-white hover:bg-slate-50 text-xs">»</button>
      `;
    }

    function changePage(p) {
      currentPage = p;
      renderDataTable();
    }

    function exportCurrentTable() {
      const rows = filteredRecords.map(r => ({
        'N°': r.id,
        'Date Souscription': r.dateSouscription,
        'Date Réception': r.dateReception,
        'Délai Réception (Jours)': r.delaiReceptionJours,
        'Affaire': r.affaire,
        'Type Affaire': r.typeAffaire,
        'Cédante (Groupe)': r.cedante,
        'Courtier / Apporteur': r.apporteur,
        'Pays': r.pays,
        'Branche': r.branche,
        'Capitaux 100%': r.capitaux100,
        'Primes 100%': r.primes100,
        'Engagement SCG-Ré': r.capitauxSCG,
        'Prime SCG-Ré': r.primesSCG,
        'Commission': r.commission,
        'Prime Nette': r.primesNettes,
        'État Bordereau': r.etatBordereau,
      }));

      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Production_ARO_SCG_Re');
      XLSX.writeFile(wb, 'SCG_Re_Production_ARO_Standardisee.xlsx');
      showToast('Export Excel téléchargé avec succès !');
    }

    // ==========================================
    // 5. UPLOAD & STANDARDIZATION ENGINE (CLIENT-SIDE)
    // ==========================================
    function setupUploadModal() {
      const modal = document.getElementById('upload-modal');
      const btnOpen = document.getElementById('btn-open-upload');
      const btnClose = document.getElementById('btn-close-upload');
      const dropZone = document.getElementById('drop-zone');
      const fileInput = document.getElementById('file-input');

      btnOpen.addEventListener('click', () => modal.classList.remove('hidden'));
      btnClose.addEventListener('click', () => modal.classList.add('hidden'));

      dropZone.addEventListener('click', () => fileInput.click());

      dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('border-brand-500', 'bg-brand-50/50');
      });

      dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('border-brand-500', 'bg-brand-50/50');
      });

      dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('border-brand-500', 'bg-brand-50/50');
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
          processExcelFile(e.dataTransfer.files[0]);
        }
      });

      fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
          processExcelFile(e.target.files[0]);
        }
      });
    }

    // STANDARDIZATION FUNCTIONS IN JAVASCRIPT
    function cleanBranchJS(b) {
      if (!b) return 'AUTRES';
      let up = b.toString().replace(/\\u00a0/g, ' ').replace(/\\s+/g, ' ').trim().toUpperCase();
      if (up.includes('DECENNALE') || up.includes('DÉCENNALE')) return 'RC DÉCENNALE';
      if (up.includes('INCENDIE')) return 'INCENDIE';
      if (up.includes('CIVILE')) return 'RESPONSABILITÉ CIVILE';
      if (up.includes('TECHNIQUE')) return 'RISQUES TECHNIQUES';
      if (up.includes('TRANSPORT')) return 'TRANSPORTS';
      if (up.includes('DIVERS')) return 'RISQUES DIVERS';
      if (up.includes('ACCIDENT')) return 'INDIVIDUELLE ACCIDENTS';
      if (up.includes('AUTO')) return 'AUTOMOBILE';
      if (up.includes('CAUTION')) return 'CAUTION';
      if (up.includes('ENERGIE') || up.includes('ÉNERGIE')) return 'ÉNERGIE';
      if (up.includes('PVT') || up.includes('TERRORISME') || up.includes('VIOLENCE')) return 'PVT';
      return up;
    }

    function cleanCountryJS(c) {
      if (!c) return 'NON DÉFINI';
      let up = c.toString().replace(/\\u00a0/g, ' ').replace(/\\s+/g, ' ').trim().toUpperCase();
      if (up.includes('COTE') || up.includes('CÔTE')) {
        if (up.includes('SENEGAL') || up.includes('SÉNÉGAL')) return "CÔTE D'IVOIRE / SÉNÉGAL";
        return "CÔTE D'IVOIRE";
      }
      if (up.includes('BENIN') || up.includes('BÉNIN')) return 'BÉNIN';
      if (up.includes('BURKINA')) return 'BURKINA FASO';
      if (up.includes('SENEGAL') || up.includes('SÉNÉGAL')) return 'SÉNÉGAL';
      if (up.includes('MALI')) return 'MALI';
      if (up.includes('TOGO')) return 'TOGO';
      if (up.includes('NIGER')) return 'NIGER';
      if (up.includes('GABON')) return 'GABON';
      if (up.includes('CAMEROUN')) return 'CAMEROUN';
      if (up.includes('GUINEE') || up.includes('GUINÉE')) {
        if (up.includes('EQUATORIAL') || up.includes('ÉQUATORIALE')) return 'GUINÉE ÉQUATORIALE';
        return 'GUINÉE CONAKRY';
      }
      if (up.includes('CONGO')) return 'CONGO';
      if (up.includes('TUNISIE')) return 'TUNISIE';
      if (up.includes('MAROC')) return 'MAROC';
      if (up.includes('ALGERIE') || up.includes('ALGÉRIE')) return 'ALGÉRIE';
      if (up.includes('MAURITANIE')) return 'MAURITANIE';
      if (up.includes('TCHAD')) return 'TCHAD';
      if (up.includes('CENTRAFRIQUE') || up.includes('RCA')) return 'CENTRAFRIQUE';
      if (up.includes('FRANCE')) return 'FRANCE';
      if (up.includes('CAP VERT') || up.includes('CABO VERDE')) return 'CAP-VERT';
      if (up.includes('GHANA')) return 'GHANA';
      return up;
    }

    function cleanApporteurJS(app) {
      if (!app) return 'NON SPÉCIFIÉ';
      let s = app.toString().trim();
      let up = s.toUpperCase();
      if (up.includes('GLUBUS')) return 'GLOBUS RE';
      if (up.includes('PREVEX')) return 'PREVEX RE';
      if (up.includes('ATLAS REINSURANCE')) return 'ATLAS REINSURANCE CONSULTANTS';
      if (up === 'SBRE' || up === 'SB RE') return 'SB RE';
      if (up.includes('KLAPTON')) return 'KLAPTON AFRICA';
      if (up === 'SUNU GROUP') return 'SUNU RE';
      return s;
    }

    function cleanCedanteJS(ced) {
      if (!ced || ced.toString().trim() === '') return 'NON SPÉCIFIÉ';
      let s = ced.toString().trim();
      let up = s.toUpperCase();
      if (up === 'TBA') return 'TBA';
      if (up.includes(',') || (up.includes('&') && !up.includes('ALLIANZ') && !up.includes('SANLAM'))) {
        return 'CO-ASSURANCE / APÉRITEUR';
      }
      if (up.includes('SANLAMALLIANZ & AMSA') || up.includes('SANLAM ALLIANZ,AMSA')) {
        return 'CO-ASSURANCE / APÉRITEUR';
      }
      if (up.includes('SUNU')) return 'SUNU ASSURANCES';
      if (up.includes('GNA')) return 'GNA';
      if (up.includes('NSIA')) return 'NSIA';
      if (up.includes('CORIS')) return 'CORIS ASSURANCES';
      if (up.includes('SANLAM') || up.includes('ALLIANZ')) return 'SANLAM ALLIANZ';
      if (up.includes('ATLANTIQUE')) return 'ATLANTIQUE ASSURANCES';
      if (up.includes('ACTIVA')) return 'ACTIVA';
      if (up.includes('AFG')) return 'AFG';
      if (up.includes('AMSA') || up.includes('ASMA')) return 'AMSA';
      if (up.includes('VISTA')) return 'VISTA ASSURANCES';
      if (up.includes('RAYNAL')) return 'RAYNAL ASSURANCES';
      if (up.includes('FINAFRICA')) return 'FINAFRICA ASSURANCES';
      if (up.includes('LANALA')) return 'LANALA ASSURANCES';
      if (up.includes('BIAT')) return 'ASSURANCES BIAT';
      if (up.includes('ASCOMA')) return 'ASCOMA';
      if (up.includes('CARTE')) return 'CARTE ASSURANCES';
      if (up.includes('ASTREE') || up.includes('ASTRÉE')) return 'ASTRÉE ASSURANCES';
      if (up.includes('ATLANTA')) return 'ATLANTA';
      if (up.includes('SONAR')) return 'SONAR';
      if (up.includes('FASO RE')) return 'FASO RE';
      if (up.includes('LEADWAY')) return 'LEADWAY ASSURANCES';
      if (up.includes('ASKIA')) return 'ASKIA ASSURANCES';
      if (up.includes('ENERALE') || up.includes('GÉNÉRALE')) return 'GÉNÉRALE DES ASSURANCES';
      if (up.includes("L'AFRICAINE")) return "L'AFRICAINE DES ASSURANCES";
      if (up.includes('UNION DES ASSURANCES DU BURKINA') || up.includes('UAB')) return 'UAB';
      if (up.includes('ZITOUNA')) return 'ZITOUNA TAKAFUL';
      if (up.includes('BH')) return 'BH ASSURANCES';
      if (up.includes('GAT')) return 'GAT ASSURANCES';
      if (up.includes('PRUDENTIAL')) return 'PRUDENTIAL BENEFICIAL';
      if (up.includes('IMPAR')) return 'IMPAR';
      if (up.includes('COMAR')) return 'COMAR';
      if (up.includes('SERENITY')) return 'SERENITY SA';
      if (up.includes('FIDELIA')) return 'FIDELIA ASSURANCES';
      if (up.includes('PREVOYANCE')) return 'PRÉVOYANCE ASSURANCES';
      if (up.includes('GEPETROL')) return 'GEPETROL SEGUROS';
      return s;
    }

    function processExcelFile(file) {
      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          
          // Chercher la feuille principale (ex: 'BASE 2026', 'BASE', ou la 1ère feuille)
          let targetSheetName = workbook.SheetNames[0];
          for (const name of workbook.SheetNames) {
            const upN = name.toUpperCase();
            if (upN.includes('BASE') || upN.includes('PRODUCTION')) {
              targetSheetName = name;
              break;
            }
          }
          const sheet = workbook.Sheets[targetSheetName];

          // Conversion en matrice brute de lignes
          const rawMatrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
          if (!rawMatrix || rawMatrix.length === 0) {
            alert('Le fichier Excel semble vide.');
            return;
          }

          // Détection automatique intelligente de la ligne des en-têtes
          let headerIdx = -1;
          for (let i = 0; i < Math.min(50, rawMatrix.length); i++) {
            const row = rawMatrix[i];
            if (Array.isArray(row)) {
              const rowStr = row.map(c => String(c || '').toUpperCase()).join(' ');
              if (rowStr.includes('DATE DE SOUSCRIPTION') || (rowStr.includes('AFFAIRE') && (rowStr.includes('CEDANTE') || rowStr.includes('APPORTEUR')))) {
                headerIdx = i;
                break;
              }
            }
          }

          if (headerIdx === -1) {
            headerIdx = 0;
          }

          const headers = rawMatrix[headerIdx].map(h => String(h || '').trim());
          
          // Recherche insensible à la casse et à la ponctuation
          // Priorise les matchs exacts, puis les matchs par inclusion
          const findColIdx = (...candidates) => {
            const normC = candidates.map(c => c.toLowerCase().replace(/[^a-z0-9]/g, ''));
            // Phase 1: Match exact (normalisé)
            const exactIdx = headers.findIndex(h => {
              const normH = h.toLowerCase().replace(/[^a-z0-9]/g, '');
              return normC.some(c => normH === c);
            });
            if (exactIdx !== -1) return exactIdx;
            // Phase 2: Match par inclusion
            return headers.findIndex(h => {
              const normH = h.toLowerCase().replace(/[^a-z0-9]/g, '');
              return normC.some(c => normH.includes(c) || c.includes(normH));
            });
          };

          const colDateSouscr = findColIdx('date de souscription', 'datesouscription');
          const colDateRecept = findColIdx('date de reception du bordereau', 'datereception', 'date de reception');
          const colAffaire = findColIdx('affaires', 'affaire');
          const colTypeAffaire = findColIdx("type d'affaire", "types d'affaires", 'typeaffaire');
          const colCedante = findColIdx('cedante', 'cedantes');
          const colApporteur = findColIdx('apporteur', 'apporteurs');
          const colPays = findColIdx('pays (lieu du risque)', 'pays', 'lieu du risque');
          const colBranche = findColIdx('branches', 'branche');
          const colCap100 = findColIdx('capitaux 100 %', 'capitaux 100%');
          const colPrim100 = findColIdx('primes 100 %', 'primes 100%');
          const colCapSCG = findColIdx('capitaux scg-re', 'capitaux scg');
          const colPrimSCG = findColIdx('primes scg-re', 'primes scg');
          const colComm = findColIdx('commission');
          const colPrimNettes = findColIdx('primes nettes', 'primes nettes de reassurance');

          const parseExcelDate = (val) => {
            if (!val) return null;
            if (val instanceof Date) return val.toISOString().split('T')[0];
            if (typeof val === 'number') {
              const d = new Date(Math.round((val - 25569) * 86400 * 1000));
              return d.toISOString().split('T')[0];
            }
            return String(val).trim();
          };

          const parseNum = (v) => {
            if (!v) return 0;
            if (typeof v === 'number') return v;
            const cleaned = String(v).replace(/\\s/g, '').replace(/,/g, '.');
            const p = parseFloat(cleaned);
            return isNaN(p) ? 0 : p;
          };

          let newRecords = [];
          let counter = 1;

          for (let r = headerIdx + 1; r < rawMatrix.length; r++) {
            const row = rawMatrix[r];
            if (!row || !Array.isArray(row)) continue;

            const affaireVal = colAffaire !== -1 ? String(row[colAffaire] || '').trim() : '';
            const idVal = row[0] ? String(row[0]).trim() : '';
            if (!affaireVal && !idVal) continue;

            const dSouscr = colDateSouscr !== -1 ? parseExcelDate(row[colDateSouscr]) : null;
            const dRecept = colDateRecept !== -1 ? parseExcelDate(row[colDateRecept]) : null;
            let delai = null;
            if (dSouscr && dRecept) {
              const t1 = new Date(dSouscr).getTime();
              const t2 = new Date(dRecept).getTime();
              if (!isNaN(t1) && !isNaN(t2)) delai = Math.round((t2 - t1) / (1000 * 3600 * 24));
            }

            const primSCG = colPrimSCG !== -1 ? parseNum(row[colPrimSCG]) : 0;
            const comm = colComm !== -1 ? parseNum(row[colComm]) : 0;
            let primNet = colPrimNettes !== -1 ? parseNum(row[colPrimNettes]) : 0;
            if (primNet === 0 && primSCG > 0) primNet = primSCG - comm;

            newRecords.push({
              id: counter++,
              dateSouscription: dSouscr,
              dateReception: dRecept,
              delaiReceptionJours: delai,
              affaire: affaireVal,
              typeAffaire: colTypeAffaire !== -1 ? String(row[colTypeAffaire] || '').trim() : '',
              cedante: cleanCedanteJS(colCedante !== -1 ? row[colCedante] : ''),
              apporteur: cleanApporteurJS(colApporteur !== -1 ? row[colApporteur] : ''),
              pays: cleanCountryJS(colPays !== -1 ? row[colPays] : ''),
              branche: cleanBranchJS(colBranche !== -1 ? row[colBranche] : ''),
              capitaux100: colCap100 !== -1 ? parseNum(row[colCap100]) : 0,
              primes100: colPrim100 !== -1 ? parseNum(row[colPrim100]) : 0,
              capitauxSCG: colCapSCG !== -1 ? parseNum(row[colCapSCG]) : 0,
              primesSCG: primSCG,
              commission: comm,
              primesNettes: primNet,
              etatBordereau: dRecept ? 'REÇU' : 'EN ATTENTE'
            });
          }

          if (newRecords.length === 0) {
            alert("Aucune ligne de production valide n'a pu être extraite.");
            return;
          }

          allRecords = newRecords;
          filteredRecords = [...allRecords];
          populateFilterOptions();
          document.getElementById('filter-branche').value = 'ALL';
          applyFilters();

          document.getElementById('upload-modal').classList.add('hidden');
          showToast(`Importation réussie : ${allRecords.length} affaires chargées et standardisées !`);
        } catch (err) {
          console.error(err);
          alert("Erreur lors de l'analyse du fichier Excel : " + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    }

  </script>
</body>
</html>
"""

# Replace logo source (base64 if available, fallback to logo_scg_re.png)
logo_src = logo_b64 if logo_b64 else "logo_scg_re.png"

final_html = html_template.replace('__INITIAL_RECORDS__', records_json).replace('__INITIAL_CLEANING_LOG__', log_json).replace('__LOGO_SRC__', logo_src)

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(final_html)

print("index.html generated successfully with SCG-Ré Logo and ARO title!")
