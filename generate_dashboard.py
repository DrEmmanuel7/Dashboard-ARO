import json
import os
import re
import base64
from datetime import datetime

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

      <!-- Action Buttons & Database Status Indicator -->
      <div class="flex items-center gap-2.5 no-print">
        <!-- Active Database Status Indicator -->
        <div id="dataset-status-pill" class="flex items-center gap-2 bg-brand-900/90 px-3.5 py-1.5 rounded-xl border border-brand-700/60 shadow-xs text-xs backdrop-blur-sm">
          <span id="dataset-status-dot" class="w-2.5 h-2.5 rounded-full bg-slate-400 ring-2 ring-white/20"></span>
          <span class="text-brand-300 font-medium text-[11px]">Base :</span>
          <span id="dataset-status-name" class="font-bold text-white max-w-[170px] truncate" title="Base d'origine SCG-Ré">Base d'origine</span>
          <span id="dataset-status-detail" class="text-brand-200/70 text-[10px] hidden md:inline">(498 affaires)</span>
          <button id="btn-reset-dataset" title="Rétablir la base de données d'origine de la SCG-Ré" class="hidden text-amber-300 hover:text-amber-100 text-[11px] font-semibold flex items-center gap-1 px-2 py-0.5 rounded bg-amber-500/20 hover:bg-amber-500/30 border border-amber-500/40 ml-1 transition-all">
            <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
            Rétablir
          </button>
        </div>

        <button id="btn-open-upload" class="inline-flex items-center gap-2 px-3.5 py-1.5 bg-gradient-to-r from-brand-600 to-brand-700 hover:from-brand-500 hover:to-brand-600 text-white rounded-xl text-xs font-semibold shadow-xs hover:shadow transition-all duration-150 border border-brand-400/30 active:scale-95">
          <svg class="w-4 h-4 text-brand-200" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>
          Importer Nouveau Fichier Excel
        </button>

        <button id="btn-download-html" title="Télécharger le fichier index.html autonome avec ces nouvelles données" class="hidden inline-flex items-center gap-2 px-3.5 py-1.5 bg-gradient-to-r from-indigo-600 to-indigo-700 hover:from-indigo-500 hover:to-indigo-600 text-white rounded-xl text-xs font-semibold shadow-xs hover:shadow transition-all duration-150 border border-indigo-400/30 active:scale-95">
          <svg class="w-4 h-4 text-indigo-200" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/></svg>
          Télécharger HTML à jour
        </button>

        <button id="btn-export-cleaned" class="inline-flex items-center gap-2 px-3.5 py-1.5 bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-500 hover:to-teal-600 text-white rounded-xl text-xs font-semibold shadow-xs hover:shadow transition-all duration-150 border border-emerald-400/30 active:scale-95">
          <svg class="w-4 h-4 text-emerald-200" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
          Exporter Excel Propre
        </button>

        <button onclick="window.print()" class="inline-flex items-center gap-2 px-3.5 py-1.5 bg-slate-800/90 hover:bg-slate-700 text-slate-200 hover:text-white rounded-xl text-xs font-semibold border border-slate-700 shadow-xs transition-all duration-150 active:scale-95">
          <svg class="w-4 h-4 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"/></svg>
          Imprimer / PDF
        </button>
      </div>

    </div>
  </header>

  <!-- STICKY FILTER BAR (PLEINE LARGEUR - FILTRE UNIQUE BRANCHE) -->
  <section class="bg-white border-b border-slate-200/90 shadow-xs sticky top-[65px] z-30 no-print">
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
          <button id="btn-reset-filters" title="Réinitialiser le filtre" class="inline-flex items-center justify-center p-2 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg border border-slate-300 shadow-xs transition-all active:scale-90">
            <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
          </button>
        </div>

        <!-- Active filter status pill & Monnaie -->
        <div class="flex items-center gap-4 text-xs text-slate-500">
          <div id="filter-status-text" class="font-medium text-slate-600 bg-slate-100 px-3.5 py-1.5 rounded-full border border-slate-200/80 shadow-xs">
            Affichage de <span class="font-bold text-brand-600" id="filtered-count">0</span> affaires sur <span class="font-bold" id="total-count-badge">0</span>
          </div>
          <div class="text-[11px] text-slate-400 font-medium">
            Monnaie de référence : <strong class="text-slate-700 font-semibold">FCFA (XOF)</strong>
          </div>
        </div>

      </div>
    </div>
  </section>

  <!-- NAVIGATION TABS (EXECUTIVE SEGMENTED CONTROL) -->
  <nav class="bg-slate-100/80 border-b border-slate-200/90 py-2.5 no-print">
    <div class="w-full px-4 sm:px-6 lg:px-8">
      <div class="inline-flex p-1 bg-white/90 rounded-xl border border-slate-200/90 shadow-xs space-x-1.5 overflow-x-auto custom-scrollbar" id="tab-nav-container">
        
        <!-- Tab 1: Souscrite -->
        <button data-tab="tab-souscrite" id="tab-btn-souscrite" class="tab-btn active px-4 py-2 rounded-lg text-xs font-bold transition-all duration-200 flex items-center gap-2 whitespace-nowrap bg-brand-600 text-white shadow-sm shadow-brand-500/25 ring-1 ring-brand-700">
          <svg class="w-4 h-4 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/></svg>
          <span>1. Vue Prime souscrite</span>
          <span id="badge-tab-souscrite" class="text-[10px] px-2 py-0.5 rounded-full font-bold bg-white/20 text-white ml-1">
            498 affaires
          </span>
        </button>

        <!-- Tab 2: Signée (Bouton Vert de base distinctif) -->
        <button data-tab="tab-signee" id="tab-btn-signee" class="tab-btn px-4 py-2 rounded-lg text-xs font-bold bg-emerald-50 text-emerald-800 border border-emerald-300/80 hover:bg-emerald-100 hover:border-emerald-400 transition-all duration-200 flex items-center gap-2 whitespace-nowrap shadow-2xs">
          <svg class="w-4 h-4 text-emerald-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
          <span>2. Vue Prime signée (Bordereaux reçus)</span>
          <span id="badge-tab-signee" class="text-[10px] px-2 py-0.5 rounded-full font-bold bg-emerald-100 text-emerald-800 border border-emerald-300/70 ml-1">
            133 bordereaux
          </span>
        </button>

        <!-- Tab 3: Explorateur -->
        <button data-tab="tab-data" id="tab-btn-data" class="tab-btn px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-50 transition-all duration-200 flex items-center gap-2 whitespace-nowrap">
          <svg class="w-4 h-4 text-slate-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h18M3 14h18m-9-4v8m-7 4h14a2 2 0 002-2V6a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>
          <span>3. Explorateur de Données & Export</span>
          <span id="badge-tab-data" class="text-[10px] px-2 py-0.5 rounded-full font-bold bg-slate-200 text-slate-700 ml-1">
            Tableau
          </span>
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

      <!-- 2x2 Grid : Branche & Pays (avec Bascule Graphique / Tableau Complet) -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Bloc 1: Prime souscrite par Branche -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-brand-500"></span>
                Prime souscrite par Branche
              </h3>
              <p class="text-[11px] text-slate-500">Classée par ordre décroissant (FCFA)</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-brand-50 text-brand-700 border border-brand-200" id="count-branches-souscrite">11 branches</span>
              <!-- Toggle Graphique / Tableau -->
              <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
                <button id="btn-toggle-souscrite-branche-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
                  Graphique Top
                </button>
                <button id="btn-toggle-souscrite-branche-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                  Tableau Complet
                </button>
              </div>
            </div>
          </div>
          <!-- Vue Graphique -->
          <div id="view-souscrite-branche-chart" class="h-80 relative flex-1">
            <canvas id="chart-souscrite-branche"></canvas>
          </div>
          <!-- Vue Tableau Complet -->
          <div id="view-souscrite-branche-table" class="hidden flex-1 overflow-y-auto custom-scrollbar max-h-80">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-50 sticky top-0 text-[11px] uppercase tracking-wider text-slate-600 border-b border-slate-200">
                <tr>
                  <th class="py-2.5 px-3 font-bold w-12 text-center">Rang</th>
                  <th class="py-2.5 px-3 font-bold">Branche</th>
                  <th class="py-2.5 px-3 font-bold text-right">Prime Souscrite</th>
                  <th class="py-2.5 px-3 font-bold text-center w-24">Affaires</th>
                  <th class="py-2.5 px-3 font-bold text-right w-28">Part %</th>
                </tr>
              </thead>
              <tbody id="table-souscrite-branche-body" class="divide-y divide-slate-100">
                <!-- Rendu Dynamique -->
              </tbody>
            </table>
          </div>
        </div>

        <!-- Bloc 2: Prime souscrite par Pays -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span>
                Prime souscrite par Pays
              </h3>
              <p class="text-[11px] text-slate-500">Lieu du risque — Classé par ordre décroissant (FCFA)</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200" id="count-pays-souscrite">0 pays</span>
              <!-- Toggle Graphique / Tableau -->
              <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
                <button id="btn-toggle-souscrite-pays-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
                  Graphique Top
                </button>
                <button id="btn-toggle-souscrite-pays-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                  Tableau Complet
                </button>
              </div>
            </div>
          </div>
          <!-- Vue Graphique -->
          <div id="view-souscrite-pays-chart" class="h-80 relative flex-1">
            <canvas id="chart-souscrite-pays"></canvas>
          </div>
          <!-- Vue Tableau Complet -->
          <div id="view-souscrite-pays-table" class="hidden flex-1 overflow-y-auto custom-scrollbar max-h-80">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-50 sticky top-0 text-[11px] uppercase tracking-wider text-slate-600 border-b border-slate-200">
                <tr>
                  <th class="py-2.5 px-3 font-bold w-12 text-center">Rang</th>
                  <th class="py-2.5 px-3 font-bold">Pays</th>
                  <th class="py-2.5 px-3 font-bold text-right">Prime Souscrite</th>
                  <th class="py-2.5 px-3 font-bold text-center w-24">Affaires</th>
                  <th class="py-2.5 px-3 font-bold text-right w-28">Part %</th>
                </tr>
              </thead>
              <tbody id="table-souscrite-pays-body" class="divide-y divide-slate-100">
                <!-- Rendu Dynamique -->
              </tbody>
            </table>
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
              <button id="btn-toggle-souscrite-cedante-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
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
                  <th class="py-2.5 px-3 font-bold w-12 text-center">Rang</th>
                  <th class="py-2.5 px-3 font-bold">Cédante (Groupe)</th>
                  <th class="py-2.5 px-3 font-bold text-right">Prime SCG-Ré</th>
                  <th class="py-2.5 px-3 font-bold text-center w-24">Affaires</th>
                  <th class="py-2.5 px-3 font-bold text-right w-28">Part %</th>
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
              <button id="btn-toggle-souscrite-courtier-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
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
                  <th class="py-2.5 px-3 font-bold w-12 text-center">Rang</th>
                  <th class="py-2.5 px-3 font-bold">Courtier / Apporteur</th>
                  <th class="py-2.5 px-3 font-bold text-right">Prime SCG-Ré</th>
                  <th class="py-2.5 px-3 font-bold text-center w-24">Affaires</th>
                  <th class="py-2.5 px-3 font-bold text-right w-28">Part %</th>
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

      <!-- 2x2 Grid : Branche & Pays (Signée avec Bascule Graphique / Tableau Complet) -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Bloc 1: Prime signée par Branche -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span>
                Prime signée par Branche
              </h3>
              <p class="text-[11px] text-slate-500">Bordereaux reçus — Tri décroissant (FCFA)</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200" id="count-branches-signee">11 branches</span>
              <!-- Toggle Graphique / Tableau -->
              <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
                <button id="btn-toggle-signee-branche-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
                  Graphique Top
                </button>
                <button id="btn-toggle-signee-branche-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                  Tableau Complet
                </button>
              </div>
            </div>
          </div>
          <!-- Vue Graphique -->
          <div id="view-signee-branche-chart" class="h-80 relative flex-1">
            <canvas id="chart-signee-branche"></canvas>
          </div>
          <!-- Vue Tableau Complet -->
          <div id="view-signee-branche-table" class="hidden flex-1 overflow-y-auto custom-scrollbar max-h-80">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-50 sticky top-0 text-[11px] uppercase tracking-wider text-slate-600 border-b border-slate-200">
                <tr>
                  <th class="py-2.5 px-3 font-bold w-12 text-center">Rang</th>
                  <th class="py-2.5 px-3 font-bold">Branche</th>
                  <th class="py-2.5 px-3 font-bold text-right">Prime Signée</th>
                  <th class="py-2.5 px-3 font-bold text-center w-24">Bordereaux</th>
                  <th class="py-2.5 px-3 font-bold text-right w-28">Part %</th>
                </tr>
              </thead>
              <tbody id="table-signee-branche-body" class="divide-y divide-slate-100">
                <!-- Rendu Dynamique -->
              </tbody>
            </table>
          </div>
        </div>

        <!-- Bloc 2: Prime signée par Pays -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-teal-500"></span>
                Prime signée par Pays
              </h3>
              <p class="text-[11px] text-slate-500">Lieu du risque — Tri décroissant (FCFA)</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-teal-50 text-teal-700 border border-teal-200" id="count-pays-signee">0 pays</span>
              <!-- Toggle Graphique / Tableau -->
              <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
                <button id="btn-toggle-signee-pays-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
                  Graphique Top
                </button>
                <button id="btn-toggle-signee-pays-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                  Tableau Complet
                </button>
              </div>
            </div>
          </div>
          <!-- Vue Graphique -->
          <div id="view-signee-pays-chart" class="h-80 relative flex-1">
            <canvas id="chart-signee-pays"></canvas>
          </div>
          <!-- Vue Tableau Complet -->
          <div id="view-signee-pays-table" class="hidden flex-1 overflow-y-auto custom-scrollbar max-h-80">
            <table class="w-full text-left text-xs border-collapse">
              <thead class="bg-slate-50 sticky top-0 text-[11px] uppercase tracking-wider text-slate-600 border-b border-slate-200">
                <tr>
                  <th class="py-2.5 px-3 font-bold w-12 text-center">Rang</th>
                  <th class="py-2.5 px-3 font-bold">Pays</th>
                  <th class="py-2.5 px-3 font-bold text-right">Prime Signée</th>
                  <th class="py-2.5 px-3 font-bold text-center w-24">Bordereaux</th>
                  <th class="py-2.5 px-3 font-bold text-right w-28">Part %</th>
                </tr>
              </thead>
              <tbody id="table-signee-pays-body" class="divide-y divide-slate-100">
                <!-- Rendu Dynamique -->
              </tbody>
            </table>
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
              <button id="btn-toggle-signee-cedante-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
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
                  <th class="py-2.5 px-3 font-bold w-12 text-center">Rang</th>
                  <th class="py-2.5 px-3 font-bold">Cédante (Groupe)</th>
                  <th class="py-2.5 px-3 font-bold text-right">Prime Signée</th>
                  <th class="py-2.5 px-3 font-bold text-center w-24">Bordereaux</th>
                  <th class="py-2.5 px-3 font-bold text-right w-28">Part %</th>
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
              <button id="btn-toggle-signee-courtier-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
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
                  <th class="py-2.5 px-3 font-bold w-12 text-center">Rang</th>
                  <th class="py-2.5 px-3 font-bold">Courtier / Apporteur</th>
                  <th class="py-2.5 px-3 font-bold text-right">Prime Signée</th>
                  <th class="py-2.5 px-3 font-bold text-center w-24">Bordereaux</th>
                  <th class="py-2.5 px-3 font-bold text-right w-28">Part %</th>
                </tr>
              </thead>
              <tbody id="table-signee-courtier-body" class="divide-y divide-slate-100">
                <!-- Rendu Dynamique -->
              </tbody>
            </table>
          </div>
        </div>

      </div>

      <!-- ========================================== -->
      <!-- BLOC 5: TOP 10 DES AFFAIRES SIGNÉES (PLEINE LARGEUR) -->
      <!-- ========================================== -->
      <div class="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden mt-2">
        <div class="bg-gradient-to-r from-emerald-950 via-emerald-900 to-teal-950 text-white p-4 flex flex-wrap items-center justify-between gap-3 border-b border-emerald-800/60">
          <div class="flex items-center gap-3">
            <span class="w-8 h-8 rounded-lg bg-emerald-800/80 border border-emerald-600/40 flex items-center justify-center text-emerald-300 font-extrabold text-sm shadow-xs">
              🏆
            </span>
            <div>
              <h3 class="text-sm font-bold text-white flex items-center gap-2">
                Top 10 des Affaires Signées (Bordereaux Reçus)
                <span class="bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 text-[10px] font-semibold px-2 py-0.5 rounded-full uppercase tracking-wider">
                  Majeures SCG-Ré
                </span>
              </h3>
              <p class="text-xs text-emerald-200/80 mt-0.5">
                Affaires fermes avec bordereaux reçus classées par Prime SCG-Ré strictement décroissante
              </p>
            </div>
          </div>
          <div class="flex items-center gap-2">
            <span id="badge-top10-signee-total" class="bg-white/10 text-emerald-100 text-xs px-3.5 py-1.5 rounded-full font-bold border border-white/20">
              Cumul Top 10 : 0 FCFA
            </span>
          </div>
        </div>

        <div class="overflow-x-auto custom-scrollbar">
          <table class="w-full text-left text-xs border-collapse">
            <thead class="bg-slate-50 text-[11px] uppercase tracking-wider text-slate-700 border-b border-slate-200 select-none">
              <tr>
                <th class="py-3 px-3 font-bold w-12 text-center">Rang</th>
                <th class="py-3 px-3 font-bold">Affaire</th>
                <th class="py-3 px-3 font-bold">Cédante</th>
                <th class="py-3 px-3 font-bold">Courtier / Apporteur</th>
                <th class="py-3 px-3 font-bold text-right">Capitaux 100 %</th>
                <th class="py-3 px-3 font-bold text-right">Prime 100 %</th>
                <th class="py-3 px-3 font-bold text-right">Taux Cession</th>
                <th class="py-3 px-3 font-bold text-right">Engagement SCG-Ré</th>
                <th class="py-3 px-3 font-bold text-right text-emerald-700">Prime SCG-Ré</th>
              </tr>
            </thead>
            <tbody id="table-signee-top10-body" class="divide-y divide-slate-100">
              <!-- Rendu Dynamique JavaScript -->
            </tbody>
          </table>
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

      <div class="mt-2.5 p-3 bg-brand-50 rounded-xl border border-brand-200 text-[11px] text-brand-900 flex items-start gap-2">
        <svg class="w-4 h-4 text-brand-600 shrink-0 mt-0.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/></svg>
        <span><strong>Mémorisation permanente :</strong> La base importée est automatiquement enregistrée dans votre navigateur. Au redémarrage ou en rouvrant le lien, <strong>vos nouvelles données restent actives par défaut</strong>. Vous pourrez rétablir la base initiale à tout moment.</span>
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
    /*__START_RECORDS__*/
    const INITIAL_RECORDS = __INITIAL_RECORDS__;
    /*__END_RECORDS__*/
    const INITIAL_CLEANING_LOG = __INITIAL_CLEANING_LOG__;
    const INITIAL_BUILD_TIMESTAMP = "__BUILD_TIMESTAMP__";

    // PERSISTENCE STORAGE CONFIG
    const DB_NAME = 'SCG_RE_ARO_DB';
    const DB_VERSION = 1;
    const STORE_NAME = 'datasets';
    const RECORD_KEY = 'latest_imported_dataset';
    const LOCAL_STORAGE_KEY = 'SCG_RE_ARO_LATEST_DATA_V2';

    // ACTIVE STATE
    let allRecords = [...INITIAL_RECORDS];
    let cleaningLog = [...INITIAL_CLEANING_LOG];
    let currentDatasetMeta = {
      isCustom: false,
      fileName: "Base d'origine SCG-Ré",
      importDate: null,
      importTimestamp: null,
      recordCount: INITIAL_RECORDS.length
    };

    // Fast synchronous pre-load via localStorage (immediate zero-latency)
    try {
      const syncItem = localStorage.getItem(LOCAL_STORAGE_KEY);
      if (syncItem) {
        const parsed = JSON.parse(syncItem);
        if (parsed && Array.isArray(parsed.records) && parsed.records.length > 0) {
          allRecords = parsed.records;
          currentDatasetMeta = {
            isCustom: true,
            fileName: parsed.fileName || 'Base personnalisée',
            importDate: parsed.importDate || null,
            importTimestamp: parsed.importTimestamp || null,
            recordCount: parsed.records.length
          };
        }
      }
    } catch (e) {
      console.warn('Sync cache check:', e);
    }

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
    // PERSISTENCE STORAGE ENGINE (IndexedDB + localStorage)
    // ==========================================
    function getIDB() {
      return new Promise((resolve, reject) => {
        if (!window.indexedDB) {
          return reject(new Error('IndexedDB non supporté'));
        }
        const req = window.indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains(STORE_NAME)) {
            db.createObjectStore(STORE_NAME);
          }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error || new Error('Erreur ouverture IDB'));
      });
    }

    async function saveDatasetToStorage(datasetObj) {
      // 1. Sauvegarde IndexedDB (capacité virtuellement illimitée)
      try {
        const db = await getIDB();
        await new Promise((resolve, reject) => {
          const tx = db.transaction(STORE_NAME, 'readwrite');
          const store = tx.objectStore(STORE_NAME);
          const req = store.put(datasetObj, RECORD_KEY);
          req.onsuccess = () => resolve();
          req.onerror = () => reject(req.error);
        });
      } catch (err) {
        console.warn('IndexedDB write warning:', err);
      }

      // 2. Sauvegarde localStorage rapide
      try {
        localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(datasetObj));
      } catch (err) {
        console.warn('localStorage write warning (limite de quota possible):', err);
      }
    }

    async function loadDatasetFromStorage() {
      // 1. Essai IndexedDB en priorité
      try {
        const db = await getIDB();
        const data = await new Promise((resolve, reject) => {
          const tx = db.transaction(STORE_NAME, 'readonly');
          const store = tx.objectStore(STORE_NAME);
          const req = store.get(RECORD_KEY);
          req.onsuccess = () => resolve(req.result);
          req.onerror = () => reject(req.error);
        });
        if (data && Array.isArray(data.records) && data.records.length > 0) {
          return data;
        }
      } catch (err) {
        console.warn('IndexedDB read warning:', err);
      }

      // 2. Repli localStorage
      try {
        const item = localStorage.getItem(LOCAL_STORAGE_KEY);
        if (item) {
          const parsed = JSON.parse(item);
          if (parsed && Array.isArray(parsed.records) && parsed.records.length > 0) {
            return parsed;
          }
        }
      } catch (err) {
        console.warn('localStorage read warning:', err);
      }

      return null;
    }

    async function clearDatasetFromStorage() {
      try {
        const db = await getIDB();
        await new Promise((resolve, reject) => {
          const tx = db.transaction(STORE_NAME, 'readwrite');
          const store = tx.objectStore(STORE_NAME);
          const req = store.delete(RECORD_KEY);
          req.onsuccess = () => resolve();
          req.onerror = () => reject(req.error);
        });
      } catch (err) {
        console.warn('IndexedDB delete warning:', err);
      }

      try {
        localStorage.removeItem(LOCAL_STORAGE_KEY);
      } catch (err) {
        console.warn('localStorage delete warning:', err);
      }
    }

    function updateDatasetStatusUI() {
      const statusName = document.getElementById('dataset-status-name');
      const statusDetail = document.getElementById('dataset-status-detail');
      const statusDot = document.getElementById('dataset-status-dot');
      const btnReset = document.getElementById('btn-reset-dataset');
      const btnDownloadHtml = document.getElementById('btn-download-html');
      const headerCount = document.getElementById('header-records-count');

      if (headerCount) {
        headerCount.textContent = `${allRecords.length} affaires chargées`;
      }

      if (currentDatasetMeta.isCustom) {
        if (statusName) {
          statusName.textContent = currentDatasetMeta.fileName || 'Base personnalisée';
          statusName.title = `Fichier : ${currentDatasetMeta.fileName || 'Inconnu'}\\nDate d'enregistrement : ${currentDatasetMeta.importDate || 'Récemment'}\\nAffaires : ${currentDatasetMeta.recordCount}`;
        }
        if (statusDetail) {
          statusDetail.textContent = `(Enregistrée par défaut • ${currentDatasetMeta.recordCount} aff.)`;
        }
        if (statusDot) {
          statusDot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
        }
        if (btnReset) btnReset.classList.remove('hidden');
        if (btnDownloadHtml) btnDownloadHtml.classList.remove('hidden');
      } else {
        if (statusName) {
          statusName.textContent = "Base d'origine SCG-Ré";
          statusName.title = "Données préchargées de base (498 affaires)";
        }
        if (statusDetail) {
          statusDetail.textContent = `(${allRecords.length} affaires)`;
        }
        if (statusDot) {
          statusDot.className = 'w-2 h-2 rounded-full bg-slate-400';
        }
        if (btnReset) btnReset.classList.add('hidden');
        if (btnDownloadHtml) btnDownloadHtml.classList.add('hidden');
      }
    }

    function setupStorageControls() {
      const btnReset = document.getElementById('btn-reset-dataset');
      if (btnReset) {
        btnReset.addEventListener('click', async () => {
          const confirmed = confirm("Voulez-vous rétablir la base de données d'origine de la SCG-Ré ?\\n\\nLa base importée sera supprimée de la mémoire locale et le tableau de bord utilisera à nouveau ses données d'origine.");
          if (confirmed) {
            await clearDatasetFromStorage();
            allRecords = [...INITIAL_RECORDS];
            filteredRecords = [...allRecords];
            currentDatasetMeta = {
              isCustom: false,
              fileName: "Base d'origine SCG-Ré",
              importDate: null,
              importTimestamp: null,
              recordCount: INITIAL_RECORDS.length
            };
            updateDatasetStatusUI();
            populateFilterOptions();
            document.getElementById('filter-branche').value = 'ALL';
            applyFilters();
            showToast("Base d'origine SCG-Ré rétablie avec succès !");
          }
        });
      }

      const btnDownload = document.getElementById('btn-download-html');
      if (btnDownload) {
        btnDownload.addEventListener('click', downloadStandaloneHtml);
      }
    }

    function downloadStandaloneHtml() {
      try {
        let fullHtml = '<!DOCTYPE html>\\n' + document.documentElement.outerHTML;
        const startMarker = '/*__START_RECORDS__*/';
        const endMarker = '/*__END_RECORDS__*/';
        const startIdx = fullHtml.indexOf(startMarker);
        const endIdx = fullHtml.indexOf(endMarker);

        if (startIdx !== -1 && endIdx !== -1) {
          const before = fullHtml.substring(0, startIdx + startMarker.length);
          const after = fullHtml.substring(endIdx);
          fullHtml = before + '\\n    const INITIAL_RECORDS = ' + JSON.stringify(allRecords) + ';\\n    ' + after;
        }

        const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'index.html';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast('Fichier index.html autonome téléchargé !');
      } catch (err) {
        console.error('Erreur export HTML:', err);
        alert('Erreur lors du téléchargement : ' + err.message);
      }
    }

    // ==========================================
    // INITIALIZATION ON DOM READY
    // ==========================================
    document.addEventListener('DOMContentLoaded', async () => {
      // Vérification asynchrone IndexedDB (au cas où les données y sont plus récentes/complètes)
      try {
        const idbData = await loadDatasetFromStorage();
        if (idbData && Array.isArray(idbData.records) && idbData.records.length > 0) {
          allRecords = idbData.records;
          currentDatasetMeta = {
            isCustom: true,
            fileName: idbData.fileName || 'Base personnalisée',
            importDate: idbData.importDate || null,
            importTimestamp: idbData.importTimestamp || null,
            recordCount: idbData.records.length
          };
          filteredRecords = [...allRecords];
        }
      } catch (err) {
        console.warn('IDB load check:', err);
      }

      setupTabs();
      setupFilters();
      setupUploadModal();
      setupTableControls();
      setupViewToggles();
      setupStorageControls();
      updateDatasetStatusUI();
      renderAll();
    });

    // ==========================================
    // TAB SWITCHING & DYNAMIC KPI CONTEXT
    // ==========================================
    function setupTabs() {
      const tabDefs = [
        { id: 'tab-btn-souscrite', target: 'tab-souscrite' },
        { id: 'tab-btn-signee', target: 'tab-signee' },
        { id: 'tab-btn-data', target: 'tab-data' }
      ];

      tabDefs.forEach(({ id, target }) => {
        const btn = document.getElementById(id);
        if (!btn) return;
        btn.addEventListener('click', () => {
          currentTabId = target;
          updateTabStyles();

          document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
          const activeContent = document.getElementById(target);
          if (activeContent) activeContent.classList.remove('hidden');

          // Strict context isolation : re-render KPIs and filter status
          renderKPIs();
          if (currentTabId === 'tab-signee') {
            renderTop10AffairesSignees();
          }

          // Resize active charts
          setTimeout(() => {
            Object.values(charts).forEach(c => {
              if (c && typeof c.resize === 'function') c.resize();
            });
          }, 60);
        });
      });

      updateTabStyles();
    }

    function updateTabStyles() {
      const btnSouscrite = document.getElementById('tab-btn-souscrite');
      const btnSignee = document.getElementById('tab-btn-signee');
      const btnData = document.getElementById('tab-btn-data');

      const badgeSouscrite = document.getElementById('badge-tab-souscrite');
      const badgeSignee = document.getElementById('badge-tab-signee');
      const badgeData = document.getElementById('badge-tab-data');

      // Styles de base au repos
      if (btnSouscrite) {
        btnSouscrite.className = 'tab-btn px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
      }
      // Onglet Prime signée : Vert émeraude subtil de base au repos
      if (btnSignee) {
        btnSignee.className = 'tab-btn px-4 py-2 rounded-lg text-xs font-bold bg-emerald-50 text-emerald-800 border border-emerald-300/80 hover:bg-emerald-100 hover:border-emerald-400 transition-all duration-150 flex items-center gap-2 whitespace-nowrap shadow-2xs';
      }
      if (btnData) {
        btnData.className = 'tab-btn px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
      }

      if (badgeSouscrite) badgeSouscrite.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-slate-200 text-slate-700 ml-1';
      if (badgeSignee) badgeSignee.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-emerald-100 text-emerald-800 border border-emerald-300/70 ml-1';
      if (badgeData) badgeData.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-slate-200 text-slate-700 ml-1';

      if (currentTabId === 'tab-souscrite' && btnSouscrite) {
        btnSouscrite.className = 'tab-btn active px-4 py-2 rounded-lg text-xs font-bold bg-brand-600 text-white shadow-sm shadow-brand-500/25 ring-1 ring-brand-700 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
        if (badgeSouscrite) badgeSouscrite.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-white/20 text-white ml-1';
      } else if (currentTabId === 'tab-signee' && btnSignee) {
        // Devient Vert Intense plein avec ombre lumineuse lorsqu'il est actif
        btnSignee.className = 'tab-btn active px-4 py-2 rounded-lg text-xs font-bold bg-emerald-600 text-white shadow-sm shadow-emerald-500/30 ring-1 ring-emerald-700 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
        if (badgeSignee) badgeSignee.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-white/20 text-white ml-1';
      } else if (currentTabId === 'tab-data' && btnData) {
        btnData.className = 'tab-btn active px-4 py-2 rounded-lg text-xs font-bold bg-slate-800 text-white shadow-sm shadow-slate-700/25 ring-1 ring-slate-900 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
        if (badgeData) badgeData.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-white/20 text-white ml-1';
      }
    }

    // ==========================================
    // TOGGLES GRAPHIQUE / TABLEAU COMPLET POUR TOUS LES BLOCS
    // ==========================================
    function bindChartTableToggle(chartBtnId, tableBtnId, chartViewId, tableViewId, chartKey) {
      const chartBtn = document.getElementById(chartBtnId);
      const tableBtn = document.getElementById(tableBtnId);
      const chartView = document.getElementById(chartViewId);
      const tableView = document.getElementById(tableViewId);
      if (!chartBtn || !tableBtn || !chartView || !tableView) return;

      chartBtn.addEventListener('click', () => {
        chartBtn.className = 'px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all';
        tableBtn.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        chartView.classList.remove('hidden');
        tableView.classList.add('hidden');
        if (chartKey && charts[chartKey] && typeof charts[chartKey].resize === 'function') {
          charts[chartKey].resize();
        }
      });

      tableBtn.addEventListener('click', () => {
        tableBtn.className = 'px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all';
        chartBtn.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        tableView.classList.remove('hidden');
        chartView.classList.add('hidden');
      });
    }

    function setupViewToggles() {
      // 1. Souscrite Branche
      bindChartTableToggle('btn-toggle-souscrite-branche-chart', 'btn-toggle-souscrite-branche-table', 'view-souscrite-branche-chart', 'view-souscrite-branche-table', 'chart-souscrite-branche');
      // 2. Souscrite Pays
      bindChartTableToggle('btn-toggle-souscrite-pays-chart', 'btn-toggle-souscrite-pays-table', 'view-souscrite-pays-chart', 'view-souscrite-pays-table', 'chart-souscrite-pays');
      // 3. Souscrite Cédante
      bindChartTableToggle('btn-toggle-souscrite-cedante-chart', 'btn-toggle-souscrite-cedante-table', 'view-souscrite-cedante-chart', 'view-souscrite-cedante-table', 'chart-souscrite-cedante');
      // 4. Souscrite Courtier
      bindChartTableToggle('btn-toggle-souscrite-courtier-chart', 'btn-toggle-souscrite-courtier-table', 'view-souscrite-courtier-chart', 'view-souscrite-courtier-table', 'chart-souscrite-courtier');

      // 5. Signée Branche
      bindChartTableToggle('btn-toggle-signee-branche-chart', 'btn-toggle-signee-branche-table', 'view-signee-branche-chart', 'view-signee-branche-table', 'chart-signee-branche');
      // 6. Signée Pays
      bindChartTableToggle('btn-toggle-signee-pays-chart', 'btn-toggle-signee-pays-table', 'view-signee-pays-chart', 'view-signee-pays-table', 'chart-signee-pays');
      // 7. Signée Cédante
      bindChartTableToggle('btn-toggle-signee-cedante-chart', 'btn-toggle-signee-cedante-table', 'view-signee-cedante-chart', 'view-signee-cedante-table', 'chart-signee-cedante');
      // 8. Signée Courtier
      bindChartTableToggle('btn-toggle-signee-courtier-chart', 'btn-toggle-signee-courtier-table', 'view-signee-courtier-chart', 'view-signee-courtier-table', 'chart-signee-courtier');
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
      updateTabBadges();
    }

    function updateTabBadges() {
      const signedCountAll = allRecords.filter(r => r.etatBordereau === 'REÇU').length;
      const signedCountFiltered = filteredRecords.filter(r => r.etatBordereau === 'REÇU').length;

      const badgeSouscrite = document.getElementById('badge-tab-souscrite');
      if (badgeSouscrite) {
        badgeSouscrite.textContent = filteredRecords.length + ' affaires';
      }
      const badgeSignee = document.getElementById('badge-tab-signee');
      if (badgeSignee) {
        badgeSignee.textContent = signedCountFiltered + ' bordereaux';
      }
      const badgeData = document.getElementById('badge-tab-data');
      if (badgeData) {
        badgeData.textContent = filteredRecords.length + ' polices';
      }
    }

    // ==========================================
    // 1. RENDER 5 PERMANENT KPIS (ÉTANCHÉITÉ STRICTE PRIME SIGNÉE)
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

      // Update Card 1: Capitaux
      document.getElementById('kpi-label-capitaux-100').textContent = isSigneeMode 
        ? 'Capitaux Signés (100%)' 
        : 'Capitaux Totaux (100%)';
      document.getElementById('kpi-capitaux-100').textContent = formatFCFA(totalCapitaux100);
      document.getElementById('kpi-sub-capitaux-100').textContent = isSigneeMode 
        ? activeSubset.length + ' bordereaux reçus'
        : '100% Marché';

      // Update Card 2: Engagement SCG-Ré
      document.getElementById('kpi-label-capitaux-scg').textContent = isSigneeMode 
        ? 'Engagement Signé SCG-Ré' 
        : 'Engagement SCG-Ré';
      document.getElementById('kpi-capitaux-scg').textContent = formatFCFA(totalCapitauxSCG);
      document.getElementById('kpi-part-capitaux-scg').textContent = formatPercent(partCapitauxSCG);

      // Update Card 3: Prime SCG-Ré
      document.getElementById('kpi-label-primes-scg').textContent = isSigneeMode 
        ? 'Prime Signée SCG-Ré' 
        : 'Prime SCG-Ré';
      document.getElementById('kpi-primes-scg').textContent = formatFCFA(totalPrimesSCG);
      document.getElementById('kpi-part-scg').textContent = formatPercent(partPrimesSCG);

      // Label adaptatif sous la Prime SCG
      const subPrimeLabel = document.getElementById('kpi-sub-primes-scg-label');
      if (subPrimeLabel) {
        subPrimeLabel.textContent = isSigneeMode ? 'Part du 100% signé :' : 'Part du 100% :';
      }

      // Update Card 4: Commission
      document.getElementById('kpi-label-commissions').textContent = isSigneeMode 
        ? 'Commission Signée' 
        : 'Commission';
      document.getElementById('kpi-commissions').textContent = formatFCFA(totalCommissions);
      document.getElementById('kpi-taux-comm-moyen').textContent = formatPercent(avgCommRate);

      // Update Card 5: Prime Nette
      document.getElementById('kpi-label-primes-nettes').textContent = isSigneeMode 
        ? 'Prime Nette Signée' 
        : 'Prime Nette';
      document.getElementById('kpi-primes-nettes').textContent = formatFCFA(totalPrimesNettes);
      document.getElementById('kpi-ratio-net').textContent = formatPercent(ratioNet);

      // Mise à jour de la barre d'état du filtre (100% étanche pour Signée)
      const filterStatusEl = document.getElementById('filter-status-text');
      if (filterStatusEl) {
        if (isSigneeMode) {
          const totalSignedBase = allRecords.filter(r => r.etatBordereau === 'REÇU').length;
          filterStatusEl.innerHTML = `Affichage de <span class="font-bold text-emerald-600">${activeSubset.length}</span> bordereaux reçus sur <span class="font-bold">${totalSignedBase}</span>`;
        } else {
          filterStatusEl.innerHTML = `Affichage de <span class="font-bold text-brand-600">${filteredRecords.length}</span> affaires sur <span class="font-bold">${allRecords.length}</span>`;
        }
      }

      document.getElementById('header-records-count').textContent = allRecords.length + ' affaires chargées';
    }

    // Helper d'agrégation retournant une liste d'objets { name, amount, count }
    function aggregateBy(records, keyExtractor) {
      const map = {};
      records.forEach(r => {
        const k = keyExtractor(r);
        if (!map[k]) map[k] = { amount: 0, count: 0 };
        map[k].amount += (r.primesSCG || 0);
        map[k].count += 1;
      });
      return Object.entries(map)
        .map(([name, data]) => ({ name, amount: data.amount, count: data.count }))
        .sort((a, b) => b.amount - a.amount);
    }

    // ==========================================
    // 2. VUE PRIME SOUSCRITE (VALEURS BRUTES)
    // ==========================================
    function renderVueSouscrite() {
      const records = filteredRecords;
      const totalSouscrite = records.reduce((acc, r) => acc + (r.primesSCG || 0), 0);
      document.getElementById('badge-souscrite-total').textContent = 'Total : ' + formatFCFA(totalSouscrite);

      // --- A. Prime souscrite par Branche ---
      const sortedBranches = aggregateBy(records, r => r.branche || 'AUTRES');
      document.getElementById('count-branches-souscrite').textContent = sortedBranches.length + ' branches';
      renderHorizontalBarChart('chart-souscrite-branche', {
        labels: sortedBranches.map(x => x.name),
        values: sortedBranches.map(x => x.amount),
        color: '#0284c7',
        datasetLabel: 'Prime souscrite'
      });
      renderRankedTable('table-souscrite-branche-body', sortedBranches, totalSouscrite, 'affaires', 'bg-brand-500');

      // --- B. Prime souscrite par Pays ---
      const sortedPays = aggregateBy(records, r => r.pays || 'NON DÉFINI');
      document.getElementById('count-pays-souscrite').textContent = sortedPays.length + ' pays';
      renderHorizontalBarChart('chart-souscrite-pays', {
        labels: sortedPays.map(x => x.name),
        values: sortedPays.map(x => x.amount),
        color: '#10b981',
        datasetLabel: 'Prime souscrite'
      });
      renderRankedTable('table-souscrite-pays-body', sortedPays, totalSouscrite, 'affaires', 'bg-emerald-500');

      // --- C. Prime souscrite par Cédante (Groupes standardisés) ---
      const sortedCedantes = aggregateBy(records, r => r.cedante || 'NON DÉFINI');
      const top10Cedantes = sortedCedantes.slice(0, 10);
      renderHorizontalBarChart('chart-souscrite-cedante', {
        labels: top10Cedantes.map(x => x.name),
        values: top10Cedantes.map(x => x.amount),
        color: '#d97706',
        datasetLabel: 'Prime souscrite'
      });
      renderRankedTable('table-souscrite-cedante-body', sortedCedantes, totalSouscrite, 'affaires', 'bg-amber-500');

      // --- D. Prime souscrite par Courtier / Apporteur ---
      const sortedCourtiers = aggregateBy(records, r => r.apporteur || 'NON SPÉCIFIÉ');
      const top10Courtiers = sortedCourtiers.slice(0, 10);
      renderHorizontalBarChart('chart-souscrite-courtier', {
        labels: top10Courtiers.map(x => x.name),
        values: top10Courtiers.map(x => x.amount),
        color: '#6366f1',
        datasetLabel: 'Prime souscrite'
      });
      renderRankedTable('table-souscrite-courtier-body', sortedCourtiers, totalSouscrite, 'affaires', 'bg-indigo-500');
    }

    // ==========================================
    // 3. VUE PRIME SIGNÉE (BORDEREAUX REÇUS)
    // ==========================================
    function renderVueSignee() {
      const records = filteredRecords.filter(r => r.etatBordereau === 'REÇU');
      const totalSignee = records.reduce((acc, r) => acc + (r.primesSCG || 0), 0);

      document.getElementById('badge-signee-count').textContent = records.length + ' bordereaux reçus';
      document.getElementById('badge-signee-total').textContent = 'Total : ' + formatFCFA(totalSignee);

      // --- A. Prime signée par Branche ---
      const sortedBranches = aggregateBy(records, r => r.branche || 'AUTRES');
      document.getElementById('count-branches-signee').textContent = sortedBranches.length + ' branches';
      renderHorizontalBarChart('chart-signee-branche', {
        labels: sortedBranches.map(x => x.name),
        values: sortedBranches.map(x => x.amount),
        color: '#059669',
        datasetLabel: 'Prime signée'
      });
      renderRankedTable('table-signee-branche-body', sortedBranches, totalSignee, 'bordereaux', 'bg-emerald-600');

      // --- B. Prime signée par Pays ---
      const sortedPays = aggregateBy(records, r => r.pays || 'NON DÉFINI');
      document.getElementById('count-pays-signee').textContent = sortedPays.length + ' pays';
      renderHorizontalBarChart('chart-signee-pays', {
        labels: sortedPays.map(x => x.name),
        values: sortedPays.map(x => x.amount),
        color: '#0d9488',
        datasetLabel: 'Prime signée'
      });
      renderRankedTable('table-signee-pays-body', sortedPays, totalSignee, 'bordereaux', 'bg-teal-600');

      // --- C. Prime signée par Cédante (Groupes standardisés) ---
      const sortedCedantes = aggregateBy(records, r => r.cedante || 'NON DÉFINI');
      const top10Cedantes = sortedCedantes.slice(0, 10);
      renderHorizontalBarChart('chart-signee-cedante', {
        labels: top10Cedantes.map(x => x.name),
        values: top10Cedantes.map(x => x.amount),
        color: '#0891b2',
        datasetLabel: 'Prime signée'
      });
      renderRankedTable('table-signee-cedante-body', sortedCedantes, totalSignee, 'bordereaux', 'bg-cyan-600');

      // --- D. Prime signée par Courtier / Apporteur ---
      const sortedCourtiers = aggregateBy(records, r => r.apporteur || 'NON SPÉCIFIÉ');
      const top10Courtiers = sortedCourtiers.slice(0, 10);
      renderHorizontalBarChart('chart-signee-courtier', {
        labels: top10Courtiers.map(x => x.name),
        values: top10Courtiers.map(x => x.amount),
        color: '#10b981',
        datasetLabel: 'Prime signée'
      });
      renderRankedTable('table-signee-courtier-body', sortedCourtiers, totalSignee, 'bordereaux', 'bg-emerald-600');

      // --- E. Top 10 des Affaires Signées ---
      renderTop10AffairesSignees();
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

    // Helper complet pour remplir les tableaux complets classés (Branche, Pays, Cédante, Courtier)
    function renderRankedTable(tbodyId, sortedList, total, countLabel, colorClass) {
      const tbody = document.getElementById(tbodyId);
      if (!tbody) return;

      if (!sortedList || sortedList.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="py-6 text-center text-slate-400 font-medium">Aucune donnée disponible</td></tr>';
        return;
      }

      const safeTotal = total > 0 ? total : 1;

      tbody.innerHTML = sortedList.map((item, idx) => {
        const rank = idx + 1;
        const part = item.amount / safeTotal;
        const partWidth = Math.min(100, Math.max(2, Math.round(part * 100)));

        let rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full text-slate-500 font-bold text-xs">${rank}</span>`;
        if (rank === 1) {
          rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-amber-100 text-amber-800 border border-amber-300 font-extrabold text-xs shadow-xs">1</span>`;
        } else if (rank === 2) {
          rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-slate-200 text-slate-700 border border-slate-300 font-extrabold text-xs shadow-xs">2</span>`;
        } else if (rank === 3) {
          rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-amber-700/10 text-amber-900 border border-amber-600/30 font-extrabold text-xs shadow-xs">3</span>`;
        }

        const countText = `${item.count} ${item.count > 1 ? countLabel : countLabel.replace(/x$/, '')}`;

        return `
          <tr class="hover:bg-slate-50/80 transition-colors">
            <td class="py-2.5 px-3 font-bold text-center">${rankBadge}</td>
            <td class="py-2.5 px-3 font-semibold text-slate-800">
              <div class="truncate max-w-[200px]" title="${item.name}">${item.name}</div>
            </td>
            <td class="py-2.5 px-3 font-extrabold text-right text-slate-900 whitespace-nowrap">${formatFCFA(item.amount)}</td>
            <td class="py-2.5 px-3 text-center whitespace-nowrap">
              <span class="px-2 py-0.5 rounded-full text-[11px] font-semibold bg-slate-100 text-slate-700 border border-slate-200/60">${countText}</span>
            </td>
            <td class="py-2.5 px-3 text-right whitespace-nowrap font-bold text-slate-700">
              <div class="flex items-center justify-end gap-2">
                <span>${formatPercent(part)}</span>
                <div class="w-12 bg-slate-100 rounded-full h-1.5 overflow-hidden">
                  <div class="${colorClass} h-1.5 rounded-full" style="width: ${partWidth}%"></div>
                </div>
              </div>
            </td>
          </tr>
        `;
      }).join('');
    }

    // ==========================================
    // TOP 10 DES AFFAIRES SIGNÉES (BORDEREAUX REÇUS)
    // ==========================================
    function renderTop10AffairesSignees() {
      const tbody = document.getElementById('table-signee-top10-body');
      if (!tbody) return;

      const signedRecs = filteredRecords.filter(r => r.etatBordereau === 'REÇU');
      const top10 = [...signedRecs]
        .sort((a, b) => (b.primesSCG || 0) - (a.primesSCG || 0))
        .slice(0, 10);

      const totalTop10 = top10.reduce((acc, r) => acc + (r.primesSCG || 0), 0);
      const totalAllSignee = signedRecs.reduce((acc, r) => acc + (r.primesSCG || 0), 0);
      const pctTop10 = totalAllSignee > 0 ? ((totalTop10 / totalAllSignee) * 100).toFixed(1) : 0;

      const badgeTotal = document.getElementById('badge-top10-signee-total');
      if (badgeTotal) {
        badgeTotal.textContent = 'Cumul Top 10 : ' + formatFCFA(totalTop10) + ' (' + pctTop10 + '% des signées)';
      }

      if (top10.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="py-6 text-center text-slate-400 font-medium">Aucun bordereau reçu pour ce filtre.</td></tr>';
        return;
      }

      tbody.innerHTML = top10.map((r, idx) => {
        const rank = idx + 1;
        let rankBadge = `<span class="inline-flex items-center justify-center w-6 h-6 rounded-full bg-slate-100 text-slate-600 font-bold text-xs">${rank}</span>`;
        if (rank === 1) {
          rankBadge = `<span class="inline-flex items-center justify-center w-6 h-6 rounded-full bg-amber-100 text-amber-800 border border-amber-300 font-extrabold text-xs shadow-xs" title="1er rang">🥇 1</span>`;
        } else if (rank === 2) {
          rankBadge = `<span class="inline-flex items-center justify-center w-6 h-6 rounded-full bg-slate-200 text-slate-800 border border-slate-300 font-extrabold text-xs shadow-xs" title="2e rang">🥈 2</span>`;
        } else if (rank === 3) {
          rankBadge = `<span class="inline-flex items-center justify-center w-6 h-6 rounded-full bg-amber-700/10 text-amber-900 border border-amber-600/30 font-extrabold text-xs shadow-xs" title="3e rang">🥉 3</span>`;
        }

        const tauxCessionVal = (r.tauxCession && r.tauxCession > 0) 
          ? (r.tauxCession > 1 ? r.tauxCession : r.tauxCession * 100).toFixed(2) + '%'
          : (r.primes100 > 0 ? ((r.primesSCG / r.primes100) * 100).toFixed(2) + '%' : '-');

        return `
          <tr class="hover:bg-emerald-50/40 transition-colors">
            <td class="py-3 px-3 text-center whitespace-nowrap">${rankBadge}</td>
            <td class="py-3 px-3 font-bold text-slate-900">
              <div class="truncate max-w-[260px]" title="${r.affaire || ''}">${r.affaire || 'NON SPÉCIFIÉ'}</div>
              <div class="text-[10px] font-semibold text-slate-500 mt-0.5 flex items-center gap-1.5">
                <span class="inline-block w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                <span>${r.branche || 'AUTRES'}</span>
                <span>•</span>
                <span>${r.pays || ''}</span>
              </div>
            </td>
            <td class="py-3 px-3 font-medium text-slate-700">
              <div class="truncate max-w-[170px]" title="${r.cedante || ''}">${r.cedante || 'NON SPÉCIFIÉ'}</div>
            </td>
            <td class="py-3 px-3 font-medium text-slate-700">
              <div class="truncate max-w-[170px]" title="${r.apporteur || ''}">${r.apporteur || 'NON SPÉCIFIÉ'}</div>
            </td>
            <td class="py-3 px-3 font-semibold text-right text-slate-700 whitespace-nowrap">${formatFCFA(r.capitaux100 || 0)}</td>
            <td class="py-3 px-3 font-semibold text-right text-slate-700 whitespace-nowrap">${formatFCFA(r.primes100 || 0)}</td>
            <td class="py-3 px-3 font-bold text-right text-brand-700 whitespace-nowrap">${tauxCessionVal}</td>
            <td class="py-3 px-3 font-semibold text-right text-purple-900 whitespace-nowrap">${formatFCFA(r.capitauxSCG || 0)}</td>
            <td class="py-3 px-3 font-extrabold text-right text-emerald-700 text-sm whitespace-nowrap bg-emerald-50/60">${formatFCFA(r.primesSCG || 0)}</td>
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
      reader.onload = async function(e) {
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

          const nowStr = new Date().toLocaleString('fr-FR');
          const metaObj = {
            records: allRecords,
            fileName: file.name,
            importDate: nowStr,
            importTimestamp: Date.now(),
            recordCount: allRecords.length
          };

          currentDatasetMeta = {
            isCustom: true,
            fileName: file.name,
            importDate: nowStr,
            importTimestamp: metaObj.importTimestamp,
            recordCount: allRecords.length
          };

          // Sauvegarde automatique et persistante (IndexedDB + localStorage)
          await saveDatasetToStorage(metaObj);

          updateDatasetStatusUI();
          populateFilterOptions();
          document.getElementById('filter-branche').value = 'ALL';
          applyFilters();

          // Réinitialiser le file input pour permettre la ré-importation du même fichier si modifié
          const fileInp = document.getElementById('file-input');
          if (fileInp) fileInp.value = '';

          document.getElementById('upload-modal').classList.add('hidden');
          showToast(`Importation réussie : ${allRecords.length} affaires enregistrées par défaut !`);
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

final_html = html_template.replace('__INITIAL_RECORDS__', records_json).replace('__INITIAL_CLEANING_LOG__', log_json).replace('__LOGO_SRC__', logo_src).replace('__BUILD_TIMESTAMP__', datetime.now().isoformat())

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(final_html)

print("index.html generated successfully with SCG-Ré Logo and ARO title!")
