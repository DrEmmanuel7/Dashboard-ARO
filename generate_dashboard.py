import json
import os
import re
import base64
from datetime import datetime

print("Starting dashboard generator...")

with open('cleaned_records.json', 'r', encoding='utf-8') as f:
    payload = json.load(f)

records_2026 = payload.get('records2026', [r for r in payload.get('records', []) if r.get('exercice') == 2026])
records_2025 = payload.get('records2025', [r for r in payload.get('records', []) if r.get('exercice') == 2025])

# Fallback if records2026 is empty
if not records_2026 and payload.get('records'):
    records_2026 = payload.get('records')

records_2026_json = json.dumps(records_2026, ensure_ascii=False)
records_2025_json = json.dumps(records_2025, ensure_ascii=False)
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
  <title>SCG-Ré | Tableau de Bord Production ARO</title>
  
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
              200: '#b9ddfd',
              300: '#7cc2fb',
              400: '#36a3f7',
              500: '#0c87eb',
              600: '#006bc9',
              700: '#0155a3',
              800: '#054886',
              900: '#0a3d6f',
              950: '#07264a',
            }
          },
          fontFamily: {
            sans: ['"Plus Jakarta Sans"', 'Inter', 'system-ui', '-apple-system', 'sans-serif'],
          }
        }
      }
    }
  </script>

  <!-- Google Fonts: Plus Jakarta Sans -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">

  <!-- Chart.js & Plugins -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
  <!-- SheetJS pour export et import Excel direct -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

  <style>
    body {
      font-family: 'Plus Jakarta Sans', sans-serif;
    }
    .custom-scrollbar::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    .custom-scrollbar::-webkit-scrollbar-track {
      background: #f1f5f9;
      border-radius: 4px;
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
      .shadow-xs, .shadow-sm, .shadow-md, .shadow-lg { box-shadow: none !important; }
      .border { border-color: #e2e8f0 !important; }
    }
  </style>
</head>
<body class="h-full flex flex-col bg-slate-50 text-slate-800 antialiased selection:bg-brand-500 selection:text-white">

  <!-- TOP HEADER (PLEINE LARGEUR) -->
  <header class="bg-white border-b border-slate-200 sticky top-0 z-40 shadow-xs no-print">
    <div class="w-full px-4 sm:px-6 lg:px-8 py-3 flex flex-wrap items-center justify-between gap-4">
      
      <!-- Brand & Title -->
      <div class="flex items-center gap-3">
        <div class="h-11 w-11 rounded-xl bg-white border border-slate-200 p-1 flex items-center justify-center shadow-xs overflow-hidden shrink-0">
          <img src="__LOGO_SRC__" alt="SCG-Ré" class="h-full w-full object-contain" onerror="this.onerror=null; this.src='https://placehold.co/100x100/0284c7/ffffff?text=SCG-R%C3%A9'">
        </div>
        <div>
          <div class="flex items-center gap-2">
            <h1 class="text-lg font-extrabold tracking-tight text-slate-900">SCG-Ré</h1>
            <span class="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-bold bg-brand-100 text-brand-800 border border-brand-200">
              FAC ARO 2026
            </span>
          </div>
          <p class="text-xs text-slate-500 font-medium">
            Tableau de Bord Exécutif — Souscriptions & Cessions Commerciales
          </p>
        </div>
      </div>

      <!-- Action Buttons & Dataset status -->
      <div class="flex items-center gap-2.5 flex-wrap">
        
        <!-- Live Dataset Badge -->
        <div class="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-100 border border-slate-200 text-xs">
          <span class="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
          <span class="text-slate-600 font-medium" id="header-records-count">Chargement...</span>
        </div>

        <!-- Bouton Sauvegarde & Télécharger HTML à jour -->
        <button id="btn-download-html" title="Télécharger le fichier index.html avec vos données intégrées" class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-300 hover:bg-emerald-100 hover:border-emerald-400 transition-all shadow-2xs">
          <svg class="w-4 h-4 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>
          <span class="hidden sm:inline">Télécharger HTML à jour</span>
        </button>

        <!-- Bouton Importer Fichier Excel -->
        <button id="btn-open-upload" class="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-bold bg-brand-600 text-white hover:bg-brand-700 active:scale-95 transition-all shadow-xs">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>
          <span>Importer Excel</span>
        </button>

        <!-- Bouton Rétablir Données d'origine -->
        <button id="btn-reset-storage" title="Rétablir les données d'origine pré-intégrées" class="p-1.5 rounded-lg text-slate-500 hover:text-rose-600 hover:bg-rose-50 border border-slate-200 transition-all">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
        </button>

        <!-- Bouton Imprimer / PDF -->
        <button onclick="window.print()" class="p-1.5 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-100 border border-slate-200 transition-all" title="Imprimer ou Exporter en PDF">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"/></svg>
        </button>
      </div>

    </div>
  </header>

  <!-- STICKY FILTER BAR (PLEINE LARGEUR - FILTRE UNIQUE BRANCHE) -->
  <section class="bg-white border-b border-slate-200/90 py-2.5 px-4 sm:px-6 lg:px-8 shadow-xs sticky top-[57px] z-30 no-print">
    <div class="w-full flex flex-wrap items-center justify-between gap-3">
      
      <!-- Filtre Unique Branche -->
      <div class="flex items-center gap-3 flex-wrap">
        <div class="flex items-center gap-2">
          <span class="w-2 h-2 rounded-full bg-brand-500"></span>
          <label for="filter-branche" class="text-xs font-bold text-slate-700 uppercase tracking-wider">
            Filtrer par Branche :
          </label>
        </div>
        <div class="relative min-w-[240px]">
          <select id="filter-branche" class="w-full pl-3 pr-8 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-xs font-semibold text-slate-800 focus:ring-2 focus:ring-brand-500 focus:border-brand-500 focus:bg-white transition-all cursor-pointer">
            <option value="">Toutes les branches (Global)</option>
          </select>
        </div>
        <button id="btn-reset-filters" title="Réinitialiser le filtre" class="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-600 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 transition-all">
          <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
          <span>Effacer le filtre</span>
        </button>
      </div>

      <!-- Compteur dynamique en direct -->
      <div class="text-xs font-medium text-slate-500 flex items-center gap-2">
        <span id="filter-status-text">Affichage de <span class="font-bold text-brand-600">518</span> affaires sur <span class="font-bold">518</span></span>
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
            0 affaires
          </span>
        </button>

        <!-- Tab 2: Signée -->
        <button data-tab="tab-signee" id="tab-btn-signee" class="tab-btn px-4 py-2 rounded-lg text-xs font-bold bg-emerald-50 text-emerald-800 border border-emerald-300/80 hover:bg-emerald-100 hover:border-emerald-400 transition-all duration-200 flex items-center gap-2 whitespace-nowrap shadow-2xs">
          <svg class="w-4 h-4 text-emerald-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
          <span>2. Vue Prime signée (Bordereaux reçus)</span>
          <span id="badge-tab-signee" class="text-[10px] px-2 py-0.5 rounded-full font-bold bg-emerald-100 text-emerald-800 border border-emerald-300/70 ml-1">
            0 bordereaux
          </span>
        </button>

        <!-- Tab 3: Affaires Sans Suite (NOUVEAU) -->
        <button data-tab="tab-sans-suite" id="tab-btn-sans-suite" class="tab-btn px-4 py-2 rounded-lg text-xs font-bold bg-rose-50 text-rose-800 border border-rose-300/80 hover:bg-rose-100 hover:border-rose-400 transition-all duration-200 flex items-center gap-2 whitespace-nowrap shadow-2xs">
          <svg class="w-4 h-4 text-rose-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M18.364 18.364A9 9 0 005.636 5.636m12.728 12.728A9 9 0 015.636 5.636m12.728 12.728L5.636 5.636"/></svg>
          <span>3. Affaires Sans Suite</span>
          <span id="badge-tab-sans-suite" class="text-[10px] px-2 py-0.5 rounded-full font-bold bg-rose-100 text-rose-800 border border-rose-300/70 ml-1">
            0 sans suite
          </span>
        </button>

        <!-- Tab 4: Explorateur -->
        <button data-tab="tab-data" id="tab-btn-data" class="tab-btn px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-50 transition-all duration-200 flex items-center gap-2 whitespace-nowrap">
          <svg class="w-4 h-4 text-slate-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h18M3 14h18m-9-4v8m-7 4h14a2 2 0 002-2V6a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>
          <span>4. Explorateur de Données & Export</span>
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
    <section id="ribbon-permanent-kpis" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
      
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
        <!-- Comparatif Subline 2025 -->
        <div id="kpi-comp-capitaux-100" class="hidden mt-1.5 pt-1.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
          <span class="text-slate-500">2025 : <strong id="kpi-comp-val-capitaux-100" class="text-slate-700">0 FCFA</strong></span>
          <span id="kpi-comp-badge-capitaux-100">0%</span>
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
        <!-- Comparatif Subline 2025 -->
        <div id="kpi-comp-capitaux-scg" class="hidden mt-1.5 pt-1.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
          <span class="text-slate-500">2025 : <strong id="kpi-comp-val-capitaux-scg" class="text-slate-700">0 FCFA</strong></span>
          <span id="kpi-comp-badge-capitaux-scg">0%</span>
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
        <!-- Comparatif Subline 2025 -->
        <div id="kpi-comp-primes-scg" class="hidden mt-1.5 pt-1.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
          <span class="text-slate-500">2025 : <strong id="kpi-comp-val-primes-scg" class="text-slate-700">0 FCFA</strong></span>
          <span id="kpi-comp-badge-primes-scg">0%</span>
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
        <!-- Comparatif Subline 2025 -->
        <div id="kpi-comp-commissions" class="hidden mt-1.5 pt-1.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
          <span class="text-slate-500">2025 : <strong id="kpi-comp-val-commissions" class="text-slate-700">0 FCFA</strong></span>
          <span id="kpi-comp-badge-commissions">0%</span>
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
        <!-- Comparatif Subline 2025 -->
        <div id="kpi-comp-primes-nettes" class="hidden mt-1.5 pt-1.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
          <span class="text-slate-500">2025 : <strong id="kpi-comp-val-primes-nettes" class="text-slate-700">0 FCFA</strong></span>
          <span id="kpi-comp-badge-primes-nettes">0%</span>
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
          
          <div id="view-souscrite-branche-chart" class="h-80 w-full relative">
            <canvas id="chart-souscrite-branche"></canvas>
          </div>
          <div id="view-souscrite-branche-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Branche</th>
                  <th class="py-2 px-3 text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-3 text-center">Affaires</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-souscrite-branche-body" class="divide-y divide-slate-100 font-medium"></tbody>
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
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200" id="count-pays-souscrite">Pays</span>
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
          
          <div id="view-souscrite-pays-chart" class="h-80 w-full relative">
            <canvas id="chart-souscrite-pays"></canvas>
          </div>
          <div id="view-souscrite-pays-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Pays</th>
                  <th class="py-2 px-3 text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-3 text-center">Affaires</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-souscrite-pays-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

      </div>

      <!-- 2x2 Grid : Cédantes & Courtiers (avec Bascule Graphique / Tableau Complet) -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Bloc 3: Prime souscrite par Cédante -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
                Prime souscrite par Cédante (Groupe)
              </h3>
              <p class="text-[11px] text-slate-500">Groupes d'assurances consolidés — Ordre décroissant</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-amber-50 text-amber-700 border border-amber-200" id="count-cedantes-souscrite">Cédantes</span>
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
          </div>
          
          <div id="view-souscrite-cedante-chart" class="h-80 w-full relative">
            <canvas id="chart-souscrite-cedante"></canvas>
          </div>
          <div id="view-souscrite-cedante-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Cédante (Groupe)</th>
                  <th class="py-2 px-3 text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-3 text-center">Affaires</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-souscrite-cedante-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

        <!-- Bloc 4: Prime souscrite par Courtier -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-indigo-500"></span>
                Prime souscrite par Courtier / Apporteur
              </h3>
              <p class="text-[11px] text-slate-500">Apporteurs standardisés — Ordre décroissant</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-indigo-50 text-indigo-700 border border-indigo-200" id="count-courtiers-souscrite">Courtiers</span>
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
          </div>
          
          <div id="view-souscrite-courtier-chart" class="h-80 w-full relative">
            <canvas id="chart-souscrite-courtier"></canvas>
          </div>
          <div id="view-souscrite-courtier-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Courtier / Apporteur</th>
                  <th class="py-2 px-3 text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-3 text-center">Affaires</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-souscrite-courtier-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

      </div>

    </div>

    <!-- ========================================== -->
    <!-- TAB 2: VUE PRIME SIGNÉE (BORDEREAUX REÇUS) -->
    <!-- ========================================== -->
    <div id="tab-signee" class="tab-content hidden space-y-6">
      
      <!-- Section Banner & Comparatif Selector -->
      <div class="bg-gradient-to-r from-emerald-900 via-emerald-800 to-teal-900 text-white rounded-xl p-4 shadow-sm flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 class="text-base font-bold flex items-center gap-2">
            <span>✅</span> Vue Prime signée — Affaires avec Bordereaux Reçus
          </h2>
          <p class="text-xs text-emerald-200 mt-0.5">
            Analyse des affaires ARO fermes dont les bordereaux ont été effectivement reçus par la SCG-Ré.
          </p>
        </div>
        
        <div class="flex items-center gap-3 flex-wrap">
          
          <!-- SÉLECTEUR 2026 / COMPARATIF 2025 vs 2026 -->
          <div class="inline-flex p-1 bg-black/25 rounded-xl border border-emerald-400/30 backdrop-blur-xs">
            <button id="btn-signee-mode-2026" class="px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all bg-white text-emerald-950 shadow-sm">
              2026
            </button>
            <button id="btn-signee-mode-comparatif" class="px-3.5 py-1.5 rounded-lg text-xs font-bold text-emerald-100 hover:text-white transition-all flex items-center gap-1.5">
              <span>⚖️</span>
              <span>Comparatif (2025 vs 2026)</span>
            </button>
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
              <p class="text-[11px] text-slate-500" id="sub-signee-branche">Bordereaux reçus — Tri décroissant (FCFA)</p>
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
          
          <div id="view-signee-branche-chart" class="h-80 w-full relative">
            <canvas id="chart-signee-branche"></canvas>
          </div>
          <div id="view-signee-branche-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead id="table-signee-branche-head" class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Branche</th>
                  <th class="py-2 px-3 text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-3 text-center">Bordereaux</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-signee-branche-body" class="divide-y divide-slate-100 font-medium"></tbody>
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
              <p class="text-[11px] text-slate-500" id="sub-signee-pays">Lieu du risque — Tri décroissant (FCFA)</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-teal-50 text-teal-700 border border-teal-200" id="count-pays-signee">Pays</span>
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
          
          <div id="view-signee-pays-chart" class="h-80 w-full relative">
            <canvas id="chart-signee-pays"></canvas>
          </div>
          <div id="view-signee-pays-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead id="table-signee-pays-head" class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Pays</th>
                  <th class="py-2 px-3 text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-3 text-center">Bordereaux</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-signee-pays-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

      </div>

      <!-- 2x2 Grid : Cédantes & Courtiers (Signée avec Bascule Graphique / Tableau Complet) -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Bloc 3: Prime signée par Cédante -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-cyan-500"></span>
                Prime signée par Cédante (Groupe)
              </h3>
              <p class="text-[11px] text-slate-500" id="sub-signee-cedante">Groupes consolidés — Ordre décroissant</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-cyan-50 text-cyan-700 border border-cyan-200" id="count-cedantes-signee">Cédantes</span>
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
          </div>
          
          <div id="view-signee-cedante-chart" class="h-80 w-full relative">
            <canvas id="chart-signee-cedante"></canvas>
          </div>
          <div id="view-signee-cedante-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead id="table-signee-cedante-head" class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Cédante (Groupe)</th>
                  <th class="py-2 px-3 text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-3 text-center">Bordereaux</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-signee-cedante-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

        <!-- Bloc 4: Prime signée par Courtier -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-emerald-600"></span>
                Prime signée par Courtier / Apporteur
              </h3>
              <p class="text-[11px] text-slate-500" id="sub-signee-courtier">Apporteurs standardisés — Ordre décroissant</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200" id="count-courtiers-signee">Courtiers</span>
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
          </div>
          
          <div id="view-signee-courtier-chart" class="h-80 w-full relative">
            <canvas id="chart-signee-courtier"></canvas>
          </div>
          <div id="view-signee-courtier-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead id="table-signee-courtier-head" class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Courtier / Apporteur</th>
                  <th class="py-2 px-3 text-right">Prime SCG-Ré</th>
                  <th class="py-2 px-3 text-center">Bordereaux</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-signee-courtier-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

      </div>

      <!-- TABLEAU DÉDIÉ : TOP 10 DES AFFAIRES SIGNÉES -->
      <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs">
        <div class="flex flex-wrap items-center justify-between mb-4 pb-3 border-b border-slate-100 gap-2">
          <div>
            <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
              <span class="w-2.5 h-2.5 rounded-full bg-emerald-600"></span>
              Top 10 des Plus Importantes Affaires Signées (Bordereaux Reçus)
            </h3>
            <p class="text-[11px] text-slate-500">
              Affaires fermes avec bordereaux reçus classées par Prime SCG-Ré strictement décroissante
            </p>
          </div>
          <span class="text-xs font-bold px-3 py-1 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200" id="badge-top10-signee-total">
            Cumul Top 10 : 0 FCFA
          </span>
        </div>

        <div class="overflow-x-auto custom-scrollbar">
          <table class="w-full text-xs text-left border-collapse">
            <thead class="bg-slate-50 border-y border-slate-200 text-slate-600 font-bold">
              <tr>
                <th class="py-2.5 px-3 text-center w-12">Rang</th>
                <th class="py-2.5 px-3">Affaire</th>
                <th class="py-2.5 px-3">Cédante (Groupe)</th>
                <th class="py-2.5 px-3">Courtier</th>
                <th class="py-2.5 px-3">Branche</th>
                <th class="py-2.5 px-3 text-right">Capitaux 100%</th>
                <th class="py-2.5 px-3 text-right">Prime 100%</th>
                <th class="py-2.5 px-3 text-center">Part Cession</th>
                <th class="py-2.5 px-3 text-right font-extrabold text-emerald-700">Prime SCG-Ré</th>
              </tr>
            </thead>
            <tbody id="table-signee-top10-body" class="divide-y divide-slate-100 font-medium"></tbody>
          </table>
        </div>
      </div>

    </div>

    <!-- ========================================== -->
    <!-- TAB 3: VUE AFFAIRES SANS SUITE (EXERCICE 2026) -->
    <!-- ========================================== -->
    <div id="tab-sans-suite" class="tab-content hidden space-y-6">
      
      <!-- 4 KPI Cards spécifiques Sans Suite (En-tête au même niveau que les onglets 1 & 2) -->
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        
        <!-- Card 1: Nombre d'affaires sans suite -->
        <div class="bg-white rounded-xl border border-slate-200 p-4 shadow-xs relative overflow-hidden group hover:border-rose-300 transition-all">
          <div class="flex items-center justify-between text-slate-500 mb-1">
            <span class="text-xs font-semibold text-slate-600">Affaires Sans Suite</span>
            <span class="p-1 rounded bg-rose-50 text-rose-600 font-bold text-xs"># DOSSIERS</span>
          </div>
          <div class="text-2xl font-extrabold text-rose-900 tracking-tight" id="kpi-ss-count">0</div>
          <div class="text-[11px] text-slate-500 mt-1 flex items-center justify-between">
            <span>Déperdition (nombre) :</span>
            <span class="font-bold text-rose-600" id="kpi-ss-taux-deperdition">0%</span>
          </div>
          <div class="absolute bottom-0 left-0 right-0 h-1 bg-rose-500"></div>
        </div>

        <!-- Card 2: Primes 100% Manquées -->
        <div class="bg-white rounded-xl border border-slate-200 p-4 shadow-xs relative overflow-hidden group hover:border-amber-300 transition-all">
          <div class="flex items-center justify-between text-slate-500 mb-1">
            <span class="text-xs font-semibold text-slate-600">Primes 100% Manquées</span>
            <span class="p-1 rounded bg-amber-50 text-amber-600 font-bold text-xs">FCFA</span>
          </div>
          <div class="text-xl font-extrabold text-amber-900 tracking-tight" id="kpi-ss-primes-100">0 FCFA</div>
          <div class="text-[11px] text-slate-500 mt-1 flex items-center justify-between">
            <span>Déperdition (prime 100%) :</span>
            <span class="font-bold text-amber-700" id="kpi-ss-taux-prime">0%</span>
          </div>
          <div class="absolute bottom-0 left-0 right-0 h-1 bg-amber-500"></div>
        </div>

        <!-- Card 3: Capitaux 100% Sans Suite -->
        <div class="bg-white rounded-xl border border-slate-200 p-4 shadow-xs relative overflow-hidden group hover:border-indigo-300 transition-all">
          <div class="flex items-center justify-between text-slate-500 mb-1">
            <span class="text-xs font-semibold text-slate-600">Capitaux 100% Sans Suite</span>
            <span class="p-1 rounded bg-indigo-50 text-indigo-600 font-bold text-xs">CAPITAUX</span>
          </div>
          <div class="text-xl font-extrabold text-indigo-950 tracking-tight" id="kpi-ss-capitaux-100">0 FCFA</div>
          <div class="text-[11px] text-slate-500 mt-1 flex items-center justify-between">
            <span>Valeurs assurées non concrétisées</span>
          </div>
          <div class="absolute bottom-0 left-0 right-0 h-1 bg-indigo-500"></div>
        </div>

        <!-- Card 4: Moyenne Primes 100% par Affaire -->
        <div class="bg-white rounded-xl border border-slate-200 p-4 shadow-xs relative overflow-hidden group hover:border-purple-300 transition-all">
          <div class="flex items-center justify-between text-slate-500 mb-1">
            <span class="text-xs font-semibold text-slate-600">Prime 100% Moyenne</span>
            <span class="p-1 rounded bg-purple-50 text-purple-600 font-bold text-xs">MOYENNE</span>
          </div>
          <div class="text-xl font-extrabold text-purple-900 tracking-tight" id="kpi-ss-moyenne-prime">0 FCFA</div>
          <div class="text-[11px] text-slate-500 mt-1 flex items-center justify-between">
            <span>Par dossier sans suite</span>
          </div>
          <div class="absolute bottom-0 left-0 right-0 h-1 bg-purple-500"></div>
        </div>

      </div>

      <!-- Section Banner -->
      <div class="bg-gradient-to-r from-rose-950 via-rose-900 to-amber-950 text-white rounded-xl p-4 shadow-sm flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 class="text-base font-bold flex items-center gap-2">
            <span>🚫</span> Vue Affaires Classées Sans Suite — Production Non Concrétisée (2026)
          </h2>
          <p class="text-xs text-rose-200 mt-0.5">
            Suivi des dossiers classés sans suite pour identifier le manque à gagner, les motifs de renonciation et les opportunités perdues.
          </p>
        </div>
        <div class="flex items-center gap-2">
          <span class="bg-rose-800/80 text-rose-100 text-xs px-3 py-1 rounded-full font-semibold border border-rose-500/40" id="badge-sans-suite-count">
            0 affaires sans suite
          </span>
          <span class="bg-white/20 text-white text-xs px-3 py-1 rounded-full font-bold" id="badge-sans-suite-total">
            Primes 100% : 0 FCFA
          </span>
        </div>
      </div>

      <!-- 2x2 Grid : 4 Blocs de Ventilation Sans Suite (Branche, Pays, Cédante, Courtier) -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Bloc 1: Branche Sans Suite -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-rose-500"></span>
                Primes 100% Perdues par Branche
              </h3>
              <p class="text-[11px] text-slate-500">Classée par Primes 100% décroissantes</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-rose-50 text-rose-700 border border-rose-200" id="count-branches-ss">Branches</span>
              <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
                <button id="btn-toggle-ss-branche-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
                  Graphique Top
                </button>
                <button id="btn-toggle-ss-branche-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                  Tableau Complet
                </button>
              </div>
            </div>
          </div>
          
          <div id="view-ss-branche-chart" class="h-80 w-full relative">
            <canvas id="chart-ss-branche"></canvas>
          </div>
          <div id="view-ss-branche-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Branche</th>
                  <th class="py-2 px-3 text-right">Primes 100%</th>
                  <th class="py-2 px-3 text-center">Affaires</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-ss-branche-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

        <!-- Bloc 2: Pays Sans Suite -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-rose-600"></span>
                Primes 100% Perdues par Pays
              </h3>
              <p class="text-[11px] text-slate-500">Lieu du risque — Classé par Primes 100% décroissantes</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-rose-50 text-rose-700 border border-rose-200" id="count-pays-ss">Pays</span>
              <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
                <button id="btn-toggle-ss-pays-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
                  Graphique Top
                </button>
                <button id="btn-toggle-ss-pays-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                  Tableau Complet
                </button>
              </div>
            </div>
          </div>
          
          <div id="view-ss-pays-chart" class="h-80 w-full relative">
            <canvas id="chart-ss-pays"></canvas>
          </div>
          <div id="view-ss-pays-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Pays</th>
                  <th class="py-2 px-3 text-right">Primes 100%</th>
                  <th class="py-2 px-3 text-center">Affaires</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-ss-pays-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

      </div>

      <!-- 2x2 Grid : Cédantes & Courtiers Sans Suite -->
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        
        <!-- Bloc 3: Cédante Sans Suite -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
                Primes 100% Perdues par Cédante (Groupe)
              </h3>
              <p class="text-[11px] text-slate-500">Groupes d'assurances consolidés — Ordre décroissant</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-amber-50 text-amber-700 border border-amber-200" id="count-cedantes-ss">Cédantes</span>
              <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
                <button id="btn-toggle-ss-cedante-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
                  Graphique Top 10
                </button>
                <button id="btn-toggle-ss-cedante-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                  Tableau Complet
                </button>
              </div>
            </div>
          </div>
          
          <div id="view-ss-cedante-chart" class="h-80 w-full relative">
            <canvas id="chart-ss-cedante"></canvas>
          </div>
          <div id="view-ss-cedante-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Cédante (Groupe)</th>
                  <th class="py-2 px-3 text-right">Primes 100%</th>
                  <th class="py-2 px-3 text-center">Affaires</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-ss-cedante-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

        <!-- Bloc 4: Courtier Sans Suite -->
        <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs flex flex-col">
          <div class="flex items-center justify-between mb-4 pb-2 border-b border-slate-100">
            <div>
              <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                <span class="w-2.5 h-2.5 rounded-full bg-orange-500"></span>
                Primes 100% Perdues par Courtier / Apporteur
              </h3>
              <p class="text-[11px] text-slate-500">Apporteurs standardisés — Ordre décroissant</p>
            </div>
            <div class="flex items-center gap-2">
              <span class="text-xs font-semibold px-2 py-0.5 rounded bg-orange-50 text-orange-700 border border-orange-200" id="count-courtiers-ss">Courtiers</span>
              <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-50 text-xs">
                <button id="btn-toggle-ss-courtier-chart" class="px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
                  Graphique Top 10
                </button>
                <button id="btn-toggle-ss-courtier-table" class="px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
                  Tableau Complet
                </button>
              </div>
            </div>
          </div>
          
          <div id="view-ss-courtier-chart" class="h-80 w-full relative">
            <canvas id="chart-ss-courtier"></canvas>
          </div>
          <div id="view-ss-courtier-table" class="h-80 w-full overflow-y-auto custom-scrollbar hidden">
            <table class="w-full text-xs text-left border-collapse">
              <thead class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10">
                <tr>
                  <th class="py-2 px-3 text-center w-12">Rang</th>
                  <th class="py-2 px-3">Courtier / Apporteur</th>
                  <th class="py-2 px-3 text-right">Primes 100%</th>
                  <th class="py-2 px-3 text-center">Affaires</th>
                  <th class="py-2 px-3 text-right">Part (%)</th>
                </tr>
              </thead>
              <tbody id="table-ss-courtier-body" class="divide-y divide-slate-100 font-medium"></tbody>
            </table>
          </div>
        </div>

      </div>

      <!-- TABLEAU DÉTAILLÉ DES AFFAIRES SANS SUITE -->
      <div class="bg-white rounded-xl border border-slate-200 p-5 shadow-xs">
        <div class="flex flex-wrap items-center justify-between mb-4 pb-3 border-b border-slate-100 gap-3">
          <div>
            <h3 class="text-sm font-bold text-slate-800 flex items-center gap-2">
              <span class="w-2.5 h-2.5 rounded-full bg-rose-600"></span>
              Liste Complète des Affaires Classées Sans Suite (2026)
            </h3>
            <p class="text-[11px] text-slate-500">
              Détail des affaires abandonnées ou retirées avec montants 100% et commentaires d'origine
            </p>
          </div>
          
          <div class="flex items-center gap-3">
            <div class="relative">
              <input type="text" id="input-search-ss" placeholder="Rechercher affaire, cédante..." class="pl-8 pr-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium focus:ring-2 focus:ring-rose-500 focus:bg-white w-56">
              <svg class="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>
            </div>
            <button id="btn-export-ss" class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-rose-50 text-rose-700 border border-rose-300 hover:bg-rose-100 transition-all">
              <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
              <span>Export Excel</span>
            </button>
          </div>
        </div>

        <div class="overflow-x-auto custom-scrollbar">
          <table class="w-full text-xs text-left border-collapse">
            <thead class="bg-slate-50 border-y border-slate-200 text-slate-600 font-bold">
              <tr>
                <th class="py-2.5 px-3 text-center w-12">N° id</th>
                <th class="py-2.5 px-3">Date Souscr.</th>
                <th class="py-2.5 px-3">Affaire</th>
                <th class="py-2.5 px-3">Cédante (Groupe)</th>
                <th class="py-2.5 px-3">Courtier</th>
                <th class="py-2.5 px-3">Branche</th>
                <th class="py-2.5 px-3">Pays</th>
                <th class="py-2.5 px-3 text-right">Capitaux 100%</th>
                <th class="py-2.5 px-3 text-right font-bold text-amber-700">Primes 100%</th>
                <th class="py-2.5 px-3">Commentaire / Motif</th>
              </tr>
            </thead>
            <tbody id="table-sans-suite-detailed-body" class="divide-y divide-slate-100 font-medium"></tbody>
          </table>
        </div>
      </div>

    </div>

    <!-- ========================================== -->
    <!-- TAB 4: DATA EXPLORER & EXPORT (ONGLET 4)   -->
    <!-- ========================================== -->
    <div id="tab-data" class="tab-content hidden space-y-4">
      
      <!-- Toolbar -->
      <div class="bg-white rounded-xl border border-slate-200 p-4 shadow-xs flex flex-wrap items-center justify-between gap-3">
        
        <div class="flex items-center gap-3 flex-wrap">
          
          <!-- SÉLECTEUR EXERCICE 2026 / 2025 (SANS TOUS) -->
          <div class="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-100 text-xs">
            <button id="btn-explorer-2026" class="px-3 py-1.5 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all">
              Exercice 2026
            </button>
            <button id="btn-explorer-2025" class="px-3 py-1.5 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all">
              Exercice 2025
            </button>
          </div>

          <!-- Search Input -->
          <div class="relative min-w-[260px]">
            <input type="text" id="table-search" placeholder="Recherche rapide (affaire, cédante, pays...)" class="w-full pl-9 pr-4 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium focus:ring-2 focus:ring-brand-500 focus:bg-white transition-all">
            <svg class="w-4 h-4 text-slate-400 absolute left-3 top-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>
          </div>

          <!-- Lignes par page -->
          <div class="flex items-center gap-2 text-xs text-slate-600">
            <span class="font-medium">Afficher :</span>
            <select id="table-page-size" class="bg-slate-50 border border-slate-200 rounded-lg px-2 py-1 text-xs font-semibold focus:ring-2 focus:ring-brand-500">
              <option value="25">25</option>
              <option value="50" selected>50</option>
              <option value="100">100</option>
              <option value="500">500</option>
              <option value="10000">Tout</option>
            </select>
          </div>

        </div>

        <div class="flex items-center gap-3">
          <span class="text-xs font-semibold text-slate-500" id="table-records-found">0 polices</span>
          <button id="btn-export-excel" class="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-bold bg-emerald-600 text-white hover:bg-emerald-700 transition-all shadow-xs">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
            <span>Exporter en Excel (.xlsx)</span>
          </button>
        </div>

      </div>

      <!-- Main Table -->
      <div class="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div class="overflow-x-auto custom-scrollbar max-h-[640px]">
          <table class="w-full text-xs text-left border-collapse" id="main-data-table">
            <thead class="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-slate-600 font-bold z-10 select-none">
              <tr>
                <th class="py-2.5 px-3 cursor-pointer hover:bg-slate-200/70" data-sort="id">ID ↕</th>
                <th class="py-2.5 px-3 cursor-pointer hover:bg-slate-200/70" data-sort="dateSouscription">Date Souscr. ↕</th>
                <th class="py-2.5 px-3 cursor-pointer hover:bg-slate-200/70" data-sort="affaire">Affaire ↕</th>
                <th class="py-2.5 px-3 cursor-pointer hover:bg-slate-200/70" data-sort="cedante">Cédante (Groupe) ↕</th>
                <th class="py-2.5 px-3 cursor-pointer hover:bg-slate-200/70" data-sort="apporteur">Courtier ↕</th>
                <th class="py-2.5 px-3 cursor-pointer hover:bg-slate-200/70" data-sort="pays">Pays ↕</th>
                <th class="py-2.5 px-3 cursor-pointer hover:bg-slate-200/70" data-sort="branche">Branche ↕</th>
                <th class="py-2.5 px-3 text-right cursor-pointer hover:bg-slate-200/70" data-sort="capitauxSCG">Capitaux SCG ↕</th>
                <th class="py-2.5 px-3 text-right cursor-pointer hover:bg-slate-200/70" data-sort="primesSCG">Prime SCG ↕</th>
                <th class="py-2.5 px-3 text-right cursor-pointer hover:bg-slate-200/70" data-sort="commission">Commission ↕</th>
                <th class="py-2.5 px-3 text-right cursor-pointer hover:bg-slate-200/70" data-sort="primesNettes">Prime Nette ↕</th>
                <th class="py-2.5 px-3 text-center cursor-pointer hover:bg-slate-200/70" data-sort="etatBordereau">Statut ↕</th>
              </tr>
            </thead>
            <tbody id="data-table-body" class="divide-y divide-slate-100 font-medium"></tbody>
          </table>
        </div>

        <!-- Pagination Controls -->
        <div class="p-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between text-xs text-slate-600 gap-3">
          <div id="table-pagination-info" class="font-medium">Affichage de 1 à 50 sur 518 affaires</div>
          <div class="flex items-center gap-1" id="table-pagination-buttons">
            <button id="btn-page-prev" class="px-2.5 py-1 rounded border border-slate-200 bg-white hover:bg-slate-50 font-bold disabled:opacity-40">Précédent</button>
            <span id="table-current-page" class="px-2 font-bold text-slate-800">Page 1</span>
            <button id="btn-page-next" class="px-2.5 py-1 rounded border border-slate-200 bg-white hover:bg-slate-50 font-bold disabled:opacity-40">Suivant</button>
          </div>
        </div>

      </div>

    </div>

  </main>

  <!-- FOOTER (PLEINE LARGEUR) -->
  <footer class="bg-white border-t border-slate-200 py-4 px-4 sm:px-6 lg:px-8 text-xs text-slate-500 flex flex-wrap items-center justify-between gap-3 no-print mt-auto">
    <div class="flex items-center gap-2">
      <span class="font-bold text-slate-700">Société Commerciale Gabonaise de Réassurance (SCG-Ré)</span>
      <span>•</span>
      <span>Souscription FAC ARO (2025 - 2026)</span>
    </div>
    <div class="flex items-center gap-4 text-slate-400">
      <span>Actualisation dynamique & persistance locale autonome</span>
    </div>
  </footer>

  <!-- MODAL IMPORT EXCEL -->
  <div id="upload-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 hidden">
    <div class="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-slate-200 space-y-4">
      <div class="flex items-center justify-between pb-3 border-b border-slate-100">
        <h3 class="text-base font-bold text-slate-900 flex items-center gap-2">
          <span>📁</span> Importer un nouveau fichier Excel
        </h3>
        <button id="btn-close-upload" class="text-slate-400 hover:text-slate-700 p-1 rounded-lg">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
        </button>
      </div>

      <p class="text-xs text-slate-600 leading-relaxed">
        Glissez-déposez votre nouveau classeur Excel de production (<code class="bg-slate-100 px-1 py-0.5 rounded text-brand-700 font-semibold">.xlsx</code> ou <code class="bg-slate-100 px-1 py-0.5 rounded text-brand-700 font-semibold">.xlsm</code>).<br>
        Le fichier sera immédiatement analysé, standardisé et sauvegardé par défaut dans votre navigateur.
      </p>

      <div id="drop-zone" class="border-2 border-dashed border-brand-300 hover:border-brand-500 rounded-xl p-8 text-center bg-brand-50/40 hover:bg-brand-50/70 transition-all cursor-pointer">
        <svg class="w-10 h-10 text-brand-500 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>
        <span class="text-xs font-bold text-brand-900 block">Cliquez ou déposez votre fichier Excel ici</span>
        <span class="text-[11px] text-slate-500 mt-1 block">Compatible avec les fichiers ARO 2026 et 2025</span>
        <input type="file" id="file-input" accept=".xlsx, .xlsm" class="hidden">
      </div>

      <div class="flex items-center justify-end gap-2 pt-2">
        <button id="btn-cancel-upload" class="px-3.5 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:bg-slate-100">Annuler</button>
      </div>
    </div>
  </div>

  <!-- TOAST NOTIFICATION -->
  <div id="toast" class="fixed bottom-5 right-5 z-50 bg-slate-900 text-white px-4 py-2.5 rounded-xl shadow-lg text-xs font-medium flex items-center gap-2 transform translate-y-20 opacity-0 transition-all duration-300 pointer-events-none">
    <span id="toast-icon">✅</span>
    <span id="toast-message">Notification</span>
  </div>

  <!-- ======================================================== -->
  <!-- JAVASCRIPT APPLICATION CORE                             -->
  <!-- ======================================================== -->
  <script>
    // Register ChartDataLabels plugin globally
    Chart.register(ChartDataLabels);

    // Formatter helpers
    const formatFCFA = (val) => {
      if (val === null || val === undefined || isNaN(val)) return '0 FCFA';
      const rounded = Math.round(val);
      const str = Math.abs(rounded).toString().replace(/\\B(?=(\\d{3})+(?!\\d))/g, ' ');
      return (rounded < 0 ? '-' : '') + str + ' FCFA';
    };

    const formatPercent = (val) => {
      if (val === null || val === undefined || isNaN(val)) return '0.0%';
      return (val * 100).toFixed(1) + '%';
    };

    const formatBarValue = (val) => {
      if (!val || val === 0) return '';
      const abs = Math.abs(val);
      if (abs >= 1e9) return (val / 1e9).toFixed(2) + ' Md';
      if (abs >= 1e6) return (val / 1e6).toFixed(1) + ' M';
      if (abs >= 1e3) return (val / 1e3).toFixed(0) + ' k';
      return Math.round(val).toString();
    };

    function showToast(msg, icon = '✅') {
      const toast = document.getElementById('toast');
      document.getElementById('toast-message').textContent = msg;
      document.getElementById('toast-icon').textContent = icon;
      toast.classList.remove('translate-y-20', 'opacity-0');
      setTimeout(() => {
        toast.classList.add('translate-y-20', 'opacity-0');
      }, 3500);
    }

    // EMBEDDED SOURCE DATA
    const INITIAL_RECORDS_2026 = __INITIAL_RECORDS_2026__;
    const INITIAL_RECORDS_2025 = __INITIAL_RECORDS_2025__;
    const INITIAL_RECORDS = INITIAL_RECORDS_2026;
    const INITIAL_CLEANING_LOG = __INITIAL_CLEANING_LOG__;
    const INITIAL_BUILD_TIMESTAMP = "__BUILD_TIMESTAMP__";

    // PERSISTENCE STORAGE CONFIG
    const DB_NAME = 'SCG_RE_ARO_DB_V3';
    const DB_VERSION = 1;
    const STORE_NAME = 'datasets';
    const RECORD_KEY = 'latest_imported_dataset';
    const LOCAL_STORAGE_KEY = 'SCG_RE_ARO_LATEST_DATA_V3';

    // ACTIVE STATE
    let allRecords = [...INITIAL_RECORDS_2026];
    let records2025 = [...INITIAL_RECORDS_2025];
    let filteredRecords = [...allRecords];
    let currentTabId = 'tab-souscrite';
    let selectedBranche = '';
    let signeeViewMode = '2026'; // '2026' or 'comparatif'
    let explorerExercice = '2026'; // '2026' or '2025'

    // Table Explorer state
    let searchTerm = '';
    let sortColumn = 'id';
    let sortDirection = 1;
    let currentPage = 1;
    let pageSize = 50;

    // Sans Suite detailed table state
    let searchSS = '';

    // Chart instances storage
    const charts = {};

    // IndexedDB helper
    function getIDB() {
      return new Promise((resolve, reject) => {
        if (!window.indexedDB) {
          resolve(null);
          return;
        }
        const req = indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains(STORE_NAME)) {
            db.createObjectStore(STORE_NAME);
          }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
      });
    }

    async function saveDatasetToStorage(dataObj) {
      try {
        localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(dataObj));
      } catch (err) {
        console.warn('LocalStorage limit exceeded, relying on IndexedDB:', err);
      }
      try {
        const db = await getIDB();
        if (db) {
          const tx = db.transaction(STORE_NAME, 'readwrite');
          tx.objectStore(STORE_NAME).put(dataObj, RECORD_KEY);
        }
      } catch (err) {
        console.error('IndexedDB save error:', err);
      }
    }

    async function loadDatasetFromStorage() {
      try {
        const db = await getIDB();
        if (db) {
          return new Promise((resolve) => {
            const tx = db.transaction(STORE_NAME, 'readonly');
            const req = tx.objectStore(STORE_NAME).get(RECORD_KEY);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => resolve(null);
          });
        }
      } catch (err) {
        console.warn('IDB read failed, fallback to localStorage', err);
      }
      try {
        const raw = localStorage.getItem(LOCAL_STORAGE_KEY);
        return raw ? JSON.parse(raw) : null;
      } catch (e) {
        return null;
      }
    }

    async function clearDatasetFromStorage() {
      try {
        localStorage.removeItem(LOCAL_STORAGE_KEY);
        const db = await getIDB();
        if (db) {
          const tx = db.transaction(STORE_NAME, 'readwrite');
          tx.objectStore(STORE_NAME).delete(RECORD_KEY);
        }
      } catch (err) {
        console.error('Clear error:', err);
      }
    }

    // ==========================================
    // INITIALIZATION & TABS SETUP
    // ==========================================
    window.addEventListener('DOMContentLoaded', async () => {
      const saved = await loadDatasetFromStorage();
      if (saved && Array.isArray(saved.records2026) && saved.records2026.length > 0) {
        allRecords = saved.records2026;
        if (Array.isArray(saved.records2025)) {
          records2025 = saved.records2025;
        }
        filteredRecords = [...allRecords];
        showToast('Données sauvegardées chargées (' + allRecords.length + ' affaires 2026)');
      }

      setupTabs();
      setupSigneeComparativeToggle();
      setupExplorerExerciceToggle();
      setupViewToggles();
      setupFilters();
      setupTableControls();
      setupSansSuiteControls();
      setupUploadModal();
      setupStorageControls();

      renderAll();
    });

    function setupTabs() {
      const tabDefs = [
        { id: 'tab-btn-souscrite', target: 'tab-souscrite' },
        { id: 'tab-btn-signee', target: 'tab-signee' },
        { id: 'tab-btn-sans-suite', target: 'tab-sans-suite' },
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

          // Gestion de l'affichage du bandeau permanent de KPI de souscription/signée
          const permRibbon = document.getElementById('ribbon-permanent-kpis');
          if (permRibbon) {
            if (target === 'tab-sans-suite' || target === 'tab-data') {
              permRibbon.classList.add('hidden');
            } else {
              permRibbon.classList.remove('hidden');
            }
          }

          renderKPIs();
          if (currentTabId === 'tab-signee') {
            renderVueSignee();
          } else if (currentTabId === 'tab-sans-suite') {
            renderVueSansSuite();
          }

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
      const btnSansSuite = document.getElementById('tab-btn-sans-suite');
      const btnData = document.getElementById('tab-btn-data');

      const badgeSouscrite = document.getElementById('badge-tab-souscrite');
      const badgeSignee = document.getElementById('badge-tab-signee');
      const badgeSansSuite = document.getElementById('badge-tab-sans-suite');
      const badgeData = document.getElementById('badge-tab-data');

      // Styles de base au repos
      if (btnSouscrite) btnSouscrite.className = 'tab-btn px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
      if (btnSignee) btnSignee.className = 'tab-btn px-4 py-2 rounded-lg text-xs font-bold bg-emerald-50 text-emerald-800 border border-emerald-300/80 hover:bg-emerald-100 hover:border-emerald-400 transition-all duration-150 flex items-center gap-2 whitespace-nowrap shadow-2xs';
      if (btnSansSuite) btnSansSuite.className = 'tab-btn px-4 py-2 rounded-lg text-xs font-bold bg-rose-50 text-rose-800 border border-rose-300/80 hover:bg-rose-100 hover:border-rose-400 transition-all duration-150 flex items-center gap-2 whitespace-nowrap shadow-2xs';
      if (btnData) btnData.className = 'tab-btn px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900 hover:bg-slate-100 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';

      if (badgeSouscrite) badgeSouscrite.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-slate-200 text-slate-700 ml-1';
      if (badgeSignee) badgeSignee.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-emerald-100 text-emerald-800 border border-emerald-300/70 ml-1';
      if (badgeSansSuite) badgeSansSuite.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-rose-100 text-rose-800 border border-rose-300/70 ml-1';
      if (badgeData) badgeData.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-slate-200 text-slate-700 ml-1';

      if (currentTabId === 'tab-souscrite' && btnSouscrite) {
        btnSouscrite.className = 'tab-btn active px-4 py-2 rounded-lg text-xs font-bold bg-brand-600 text-white shadow-sm shadow-brand-500/25 ring-1 ring-brand-700 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
        if (badgeSouscrite) badgeSouscrite.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-white/20 text-white ml-1';
      } else if (currentTabId === 'tab-signee' && btnSignee) {
        btnSignee.className = 'tab-btn active px-4 py-2 rounded-lg text-xs font-bold bg-emerald-600 text-white shadow-sm shadow-emerald-500/30 ring-1 ring-emerald-700 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
        if (badgeSignee) badgeSignee.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-white/20 text-white ml-1';
      } else if (currentTabId === 'tab-sans-suite' && btnSansSuite) {
        btnSansSuite.className = 'tab-btn active px-4 py-2 rounded-lg text-xs font-bold bg-rose-600 text-white shadow-sm shadow-rose-500/30 ring-1 ring-rose-700 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
        if (badgeSansSuite) badgeSansSuite.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-white/20 text-white ml-1';
      } else if (currentTabId === 'tab-data' && btnData) {
        btnData.className = 'tab-btn active px-4 py-2 rounded-lg text-xs font-bold bg-slate-800 text-white shadow-sm shadow-slate-700/25 ring-1 ring-slate-900 transition-all duration-150 flex items-center gap-2 whitespace-nowrap';
        if (badgeData) badgeData.className = 'text-[10px] px-2 py-0.5 rounded-full font-bold bg-white/20 text-white ml-1';
      }
    }

    // Sélecteur 2026 / Comparatif dans l'onglet Prime signée
    function setupSigneeComparativeToggle() {
      const btn2026 = document.getElementById('btn-signee-mode-2026');
      const btnComp = document.getElementById('btn-signee-mode-comparatif');
      if (!btn2026 || !btnComp) return;

      btn2026.addEventListener('click', () => {
        signeeViewMode = '2026';
        btn2026.className = 'px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all bg-white text-emerald-950 shadow-sm';
        btnComp.className = 'px-3.5 py-1.5 rounded-lg text-xs font-bold text-emerald-100 hover:text-white transition-all flex items-center gap-1.5';
        renderKPIs();
        renderVueSignee();
      });

      btnComp.addEventListener('click', () => {
        signeeViewMode = 'comparatif';
        btnComp.className = 'px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all bg-white text-emerald-950 shadow-sm flex items-center gap-1.5';
        btn2026.className = 'px-3.5 py-1.5 rounded-lg text-xs font-bold text-emerald-100 hover:text-white transition-all';
        renderKPIs();
        renderVueSignee();
      });
    }

    // Sélecteur d'exercice dans l'Explorateur (2026 vs 2025)
    function setupExplorerExerciceToggle() {
      const btn2026 = document.getElementById('btn-explorer-2026');
      const btn2025 = document.getElementById('btn-explorer-2025');
      if (!btn2026 || !btn2025) return;

      btn2026.addEventListener('click', () => {
        explorerExercice = '2026';
        btn2026.className = 'px-3 py-1.5 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all';
        btn2025.className = 'px-3 py-1.5 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        currentPage = 1;
        renderDataTable();
      });

      btn2025.addEventListener('click', () => {
        explorerExercice = '2025';
        btn2025.className = 'px-3 py-1.5 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all';
        btn2026.className = 'px-3 py-1.5 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        currentPage = 1;
        renderDataTable();
      });
    }

    // ==========================================
    // TOGGLES GRAPHIQUE / TABLEAU COMPLET
    // ==========================================
    function bindChartTableToggle(btnChartId, btnTableId, viewChartId, viewTableId) {
      const btnChart = document.getElementById(btnChartId);
      const btnTable = document.getElementById(btnTableId);
      const viewChart = document.getElementById(viewChartId);
      const viewTable = document.getElementById(viewTableId);

      if (!btnChart || !btnTable || !viewChart || !viewTable) return;

      btnChart.addEventListener('click', () => {
        viewChart.classList.remove('hidden');
        viewTable.classList.add('hidden');
        btnChart.className = 'px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all';
        btnTable.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
        // Resize canvas
        const canvas = viewChart.querySelector('canvas');
        if (canvas && charts[canvas.id]) charts[canvas.id].resize();
      });

      btnTable.addEventListener('click', () => {
        viewTable.classList.remove('hidden');
        viewChart.classList.add('hidden');
        btnTable.className = 'px-2.5 py-1 rounded-md font-bold bg-white text-slate-800 shadow-xs border border-slate-200/80 transition-all';
        btnChart.className = 'px-2.5 py-1 rounded-md font-semibold text-slate-500 hover:text-slate-800 transition-all';
      });
    }

    function setupViewToggles() {
      // 1. Vue Souscrite
      bindChartTableToggle('btn-toggle-souscrite-branche-chart', 'btn-toggle-souscrite-branche-table', 'view-souscrite-branche-chart', 'view-souscrite-branche-table');
      bindChartTableToggle('btn-toggle-souscrite-pays-chart', 'btn-toggle-souscrite-pays-table', 'view-souscrite-pays-chart', 'view-souscrite-pays-table');
      bindChartTableToggle('btn-toggle-souscrite-cedante-chart', 'btn-toggle-souscrite-cedante-table', 'view-souscrite-cedante-chart', 'view-souscrite-cedante-table');
      bindChartTableToggle('btn-toggle-souscrite-courtier-chart', 'btn-toggle-souscrite-courtier-table', 'view-souscrite-courtier-chart', 'view-souscrite-courtier-table');

      // 2. Vue Signée
      bindChartTableToggle('btn-toggle-signee-branche-chart', 'btn-toggle-signee-branche-table', 'view-signee-branche-chart', 'view-signee-branche-table');
      bindChartTableToggle('btn-toggle-signee-pays-chart', 'btn-toggle-signee-pays-table', 'view-signee-pays-chart', 'view-signee-pays-table');
      bindChartTableToggle('btn-toggle-signee-cedante-chart', 'btn-toggle-signee-cedante-table', 'view-signee-cedante-chart', 'view-signee-cedante-table');
      bindChartTableToggle('btn-toggle-signee-courtier-chart', 'btn-toggle-signee-courtier-table', 'view-signee-courtier-chart', 'view-signee-courtier-table');

      // 3. Vue Sans Suite
      bindChartTableToggle('btn-toggle-ss-branche-chart', 'btn-toggle-ss-branche-table', 'view-ss-branche-chart', 'view-ss-branche-table');
      bindChartTableToggle('btn-toggle-ss-pays-chart', 'btn-toggle-ss-pays-table', 'view-ss-pays-chart', 'view-ss-pays-table');
      bindChartTableToggle('btn-toggle-ss-cedante-chart', 'btn-toggle-ss-cedante-table', 'view-ss-cedante-chart', 'view-ss-cedante-table');
      bindChartTableToggle('btn-toggle-ss-courtier-chart', 'btn-toggle-ss-courtier-table', 'view-ss-courtier-chart', 'view-ss-courtier-table');
    }

    // ==========================================
    // FILTERS
    // ==========================================
    function setupFilters() {
      const filterBranche = document.getElementById('filter-branche');
      const btnReset = document.getElementById('btn-reset-filters');

      populateFilterOptions();

      if (filterBranche) {
        filterBranche.addEventListener('change', (e) => {
          selectedBranche = e.target.value;
          applyFilters();
        });
      }

      if (btnReset) {
        btnReset.addEventListener('click', () => {
          selectedBranche = '';
          if (filterBranche) filterBranche.value = '';
          applyFilters();
          showToast('Filtre réinitialisé');
        });
      }
    }

    function populateFilterOptions() {
      const filterBranche = document.getElementById('filter-branche');
      if (!filterBranche) return;

      const branches = new Set();
      allRecords.forEach(r => {
        if (r.branche) branches.add(r.branche);
      });

      const currentVal = filterBranche.value;
      filterBranche.innerHTML = '<option value="">Toutes les branches (Global)</option>';
      Array.from(branches).sort().forEach(b => {
        const count = allRecords.filter(r => r.branche === b).length;
        filterBranche.innerHTML += `<option value="${b}">${b} (${count} affaires)</option>`;
      });
      filterBranche.value = currentVal;
    }

    function applyFilters() {
      filteredRecords = allRecords.filter(r => {
        if (selectedBranche && r.branche !== selectedBranche) return false;
        return true;
      });
      currentPage = 1;
      renderAll();
    }

    function renderAll() {
      renderKPIs();
      renderVueSouscrite();
      renderVueSignee();
      renderVueSansSuite();
      renderDataTable();
      updateTabBadges();
    }

    function updateTabBadges() {
      const signedCountFiltered = filteredRecords.filter(r => r.etatBordereau === 'REÇU').length;
      const ssCountFiltered = filteredRecords.filter(r => r.isSansSuite === true).length;

      const badgeSouscrite = document.getElementById('badge-tab-souscrite');
      if (badgeSouscrite) badgeSouscrite.textContent = filteredRecords.length + ' affaires';

      const badgeSignee = document.getElementById('badge-tab-signee');
      if (badgeSignee) badgeSignee.textContent = signedCountFiltered + ' bordereaux';

      const badgeSansSuite = document.getElementById('badge-tab-sans-suite');
      if (badgeSansSuite) badgeSansSuite.textContent = ssCountFiltered + ' sans suite';

      const badgeData = document.getElementById('badge-tab-data');
      if (badgeData) badgeData.textContent = (explorerExercice === '2026' ? filteredRecords.length : records2025.length) + ' polices';
    }

    // Helper pour générer un badge de variation (+X% vert / -X% rouge)
    function formatVariationBadge(val26, val25) {
      if (!val25 && !val26) return `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-600">0.0%</span>`;
      if (!val25) return `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">+100%</span>`;
      const pct = ((val26 - val25) / val25) * 100;
      if (Math.abs(pct) < 0.05) return `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-600">0.0%</span>`;
      const isPos = pct > 0;
      const sign = isPos ? '+' : '';
      const color = isPos ? 'bg-emerald-100 text-emerald-800 border border-emerald-200' : 'bg-rose-100 text-rose-800 border border-rose-200';
      return `<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold ${color}">${sign}${pct.toFixed(1)}%</span>`;
    }

    // ==========================================
    // 1. RENDER 5 PERMANENT KPIS
    // ==========================================
    function renderKPIs() {
      const isSigneeMode = (currentTabId === 'tab-signee');
      const isSansSuiteTab = (currentTabId === 'tab-sans-suite');
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

      // GESTION DU COMPARATIF SOUS LES KPI (QUAND EN MODE COMPARATIF DANS L'ONGLET SIGNÉE)
      const isComparatifActive = (isSigneeMode && signeeViewMode === 'comparatif');
      const compIds = [
        'kpi-comp-capitaux-100',
        'kpi-comp-capitaux-scg',
        'kpi-comp-primes-scg',
        'kpi-comp-commissions',
        'kpi-comp-primes-nettes'
      ];

      if (isComparatifActive) {
        const active2025Signed = records2025.filter(r => r.etatBordereau === 'REÇU' && (!selectedBranche || r.branche === selectedBranche));
        const c100_25 = active2025Signed.reduce((acc, r) => acc + (r.capitaux100 || 0), 0);
        const cSCG_25 = active2025Signed.reduce((acc, r) => acc + (r.capitauxSCG || 0), 0);
        const pSCG_25 = active2025Signed.reduce((acc, r) => acc + (r.primesSCG || 0), 0);
        const comm_25 = active2025Signed.reduce((acc, r) => acc + (r.commission || 0), 0);
        const pNet_25 = active2025Signed.reduce((acc, r) => acc + (r.primesNettes || 0), 0);

        compIds.forEach(id => {
          const el = document.getElementById(id);
          if (el) el.classList.remove('hidden');
        });

        document.getElementById('kpi-comp-val-capitaux-100').textContent = formatFCFA(c100_25);
        document.getElementById('kpi-comp-badge-capitaux-100').innerHTML = formatVariationBadge(totalCapitaux100, c100_25);

        document.getElementById('kpi-comp-val-capitaux-scg').textContent = formatFCFA(cSCG_25);
        document.getElementById('kpi-comp-badge-capitaux-scg').innerHTML = formatVariationBadge(totalCapitauxSCG, cSCG_25);

        document.getElementById('kpi-comp-val-primes-scg').textContent = formatFCFA(pSCG_25);
        document.getElementById('kpi-comp-badge-primes-scg').innerHTML = formatVariationBadge(totalPrimesSCG, pSCG_25);

        document.getElementById('kpi-comp-val-commissions').textContent = formatFCFA(comm_25);
        document.getElementById('kpi-comp-badge-commissions').innerHTML = formatVariationBadge(totalCommissions, comm_25);

        document.getElementById('kpi-comp-val-primes-nettes').textContent = formatFCFA(pNet_25);
        document.getElementById('kpi-comp-badge-primes-nettes').innerHTML = formatVariationBadge(totalPrimesNettes, pNet_25);
      } else {
        compIds.forEach(id => {
          const el = document.getElementById(id);
          if (el) el.classList.add('hidden');
        });
      }

      // Filter status bar text
      const filterStatusEl = document.getElementById('filter-status-text');
      if (filterStatusEl) {
        if (isSigneeMode) {
          const totalSignedBase = allRecords.filter(r => r.etatBordereau === 'REÇU').length;
          filterStatusEl.innerHTML = `Affichage de <span class="font-bold text-emerald-600">${activeSubset.length}</span> bordereaux reçus sur <span class="font-bold">${totalSignedBase}</span>`;
        } else if (isSansSuiteTab) {
          const totalSSBase = allRecords.filter(r => r.isSansSuite === true).length;
          const currentSS = filteredRecords.filter(r => r.isSansSuite === true).length;
          filterStatusEl.innerHTML = `Affichage de <span class="font-bold text-rose-600">${currentSS}</span> affaires sans suite sur <span class="font-bold">${totalSSBase}</span>`;
        } else {
          filterStatusEl.innerHTML = `Affichage de <span class="font-bold text-brand-600">${filteredRecords.length}</span> affaires sur <span class="font-bold">${allRecords.length}</span>`;
        }
      }

      document.getElementById('header-records-count').textContent = allRecords.length + ' affaires 2026 chargées';
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

    // Helper d'agrégation comparative 2025 vs 2026
    function aggregateComparative(records26, records25, keyExtractor) {
      const map = {};
      records26.forEach(r => {
        const k = keyExtractor(r);
        if (!map[k]) map[k] = { name: k, amount2026: 0, count2026: 0, amount2025: 0, count2025: 0 };
        map[k].amount2026 += (r.primesSCG || 0);
        map[k].count2026 += 1;
      });
      records25.forEach(r => {
        const k = keyExtractor(r);
        if (!map[k]) map[k] = { name: k, amount2026: 0, count2026: 0, amount2025: 0, count2025: 0 };
        map[k].amount2025 += (r.primesSCG || 0);
        map[k].count2025 += 1;
      });
      const list = Object.values(map);
      list.sort((a, b) => (b.amount2026 - a.amount2026) || (b.amount2025 - a.amount2025) || a.name.localeCompare(b.name));
      list.forEach(item => {
        item.delta = item.amount2026 - item.amount2025;
        item.growthPct = item.amount2025 > 0 ? ((item.amount2026 - item.amount2025) / item.amount2025) * 100 : (item.amount2026 > 0 ? 100 : 0);
      });
      return list;
    }

    // Helper d'agrégation par somme d'une colonne spécifique (ex: primes100 pour sans suite)
    function aggregateBySum(records, keyExtractor, valExtractor) {
      const map = {};
      records.forEach(r => {
        const k = keyExtractor(r);
        if (!map[k]) map[k] = { name: k, amount: 0, count: 0 };
        map[k].amount += valExtractor(r);
        map[k].count += 1;
      });
      return Object.values(map).sort((a, b) => b.amount - a.amount);
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
    // 3. VUE PRIME SIGNÉE (AVEC MODE COMPARATIF 2025 vs 2026)
    // ==========================================
    function renderVueSignee() {
      if (signeeViewMode === 'comparatif') {
        renderVueSigneeComparative();
      } else {
        renderVueSigneeStandard();
      }
    }

    function renderVueSigneeStandard() {
      const records = filteredRecords.filter(r => r.etatBordereau === 'REÇU');
      const totalSignee = records.reduce((acc, r) => acc + (r.primesSCG || 0), 0);

      document.getElementById('badge-signee-count').textContent = records.length + ' bordereaux reçus';
      document.getElementById('badge-signee-total').textContent = 'Total : ' + formatFCFA(totalSignee);

      document.getElementById('sub-signee-branche').textContent = 'Bordereaux reçus — Tri décroissant (FCFA)';
      document.getElementById('sub-signee-pays').textContent = 'Lieu du risque — Tri décroissant (FCFA)';
      document.getElementById('sub-signee-cedante').textContent = 'Groupes consolidés — Ordre décroissant';
      document.getElementById('sub-signee-courtier').textContent = 'Apporteurs standardisés — Ordre décroissant';

      // Reset Table Heads to Standard
      resetStandardTableHead('table-signee-branche-head', 'Branche');
      resetStandardTableHead('table-signee-pays-head', 'Pays');
      resetStandardTableHead('table-signee-cedante-head', 'Cédante (Groupe)');
      resetStandardTableHead('table-signee-courtier-head', 'Courtier / Apporteur');

      // --- A. Branche ---
      const sortedBranches = aggregateBy(records, r => r.branche || 'AUTRES');
      document.getElementById('count-branches-signee').textContent = sortedBranches.length + ' branches';
      renderHorizontalBarChart('chart-signee-branche', {
        labels: sortedBranches.map(x => x.name),
        values: sortedBranches.map(x => x.amount),
        color: '#059669',
        datasetLabel: 'Prime signée'
      });
      renderRankedTable('table-signee-branche-body', sortedBranches, totalSignee, 'bordereaux', 'bg-emerald-600');

      // --- B. Pays ---
      const sortedPays = aggregateBy(records, r => r.pays || 'NON DÉFINI');
      document.getElementById('count-pays-signee').textContent = sortedPays.length + ' pays';
      renderHorizontalBarChart('chart-signee-pays', {
        labels: sortedPays.map(x => x.name),
        values: sortedPays.map(x => x.amount),
        color: '#0d9488',
        datasetLabel: 'Prime signée'
      });
      renderRankedTable('table-signee-pays-body', sortedPays, totalSignee, 'bordereaux', 'bg-teal-600');

      // --- C. Cédante ---
      const sortedCedantes = aggregateBy(records, r => r.cedante || 'NON DÉFINI');
      const top10Cedantes = sortedCedantes.slice(0, 10);
      renderHorizontalBarChart('chart-signee-cedante', {
        labels: top10Cedantes.map(x => x.name),
        values: top10Cedantes.map(x => x.amount),
        color: '#0891b2',
        datasetLabel: 'Prime signée'
      });
      renderRankedTable('table-signee-cedante-body', sortedCedantes, totalSignee, 'bordereaux', 'bg-cyan-600');

      // --- D. Courtier ---
      const sortedCourtiers = aggregateBy(records, r => r.apporteur || 'NON SPÉCIFIÉ');
      const top10Courtiers = sortedCourtiers.slice(0, 10);
      renderHorizontalBarChart('chart-signee-courtier', {
        labels: top10Courtiers.map(x => x.name),
        values: top10Courtiers.map(x => x.amount),
        color: '#10b981',
        datasetLabel: 'Prime signée'
      });
      renderRankedTable('table-signee-courtier-body', sortedCourtiers, totalSignee, 'bordereaux', 'bg-emerald-600');

      // --- E. Top 10 Affaires Signées ---
      renderTop10AffairesSignees();
    }

    function renderVueSigneeComparative() {
      const records26 = filteredRecords.filter(r => r.etatBordereau === 'REÇU');
      const records25 = records2025.filter(r => r.etatBordereau === 'REÇU' && (!selectedBranche || r.branche === selectedBranche));

      const totalSignee26 = records26.reduce((acc, r) => acc + (r.primesSCG || 0), 0);
      const totalSignee25 = records25.reduce((acc, r) => acc + (r.primesSCG || 0), 0);
      const deltaTotal = totalSignee26 - totalSignee25;
      const pctTotal = totalSignee25 > 0 ? (deltaTotal / totalSignee25) * 100 : 0;
      const sign = pctTotal > 0 ? '+' : '';

      document.getElementById('badge-signee-count').textContent = `${records26.length} bordereaux (2026) vs ${records25.length} (2025)`;
      document.getElementById('badge-signee-total').textContent = `2026 : ${formatFCFA(totalSignee26)} | 2025 : ${formatFCFA(totalSignee25)} (${sign}${pctTotal.toFixed(1)}%)`;

      document.getElementById('sub-signee-branche').textContent = 'Comparatif 2026 vs 2025 avec variation (%)';
      document.getElementById('sub-signee-pays').textContent = 'Comparatif 2026 vs 2025 avec variation (%)';
      document.getElementById('sub-signee-cedante').textContent = 'Comparatif 2026 vs 2025 avec variation (%)';
      document.getElementById('sub-signee-courtier').textContent = 'Comparatif 2026 vs 2025 avec variation (%)';

      // --- A. Branche Comparatif ---
      const compBranches = aggregateComparative(records26, records25, r => r.branche || 'AUTRES');
      document.getElementById('count-branches-signee').textContent = compBranches.length + ' branches';
      renderComparativeHorizontalBarChart('chart-signee-branche', compBranches);
      renderComparativeRankedTable('table-signee-branche-body', 'table-signee-branche-head', compBranches, 'Branche');

      // --- B. Pays Comparatif ---
      const compPays = aggregateComparative(records26, records25, r => r.pays || 'NON DÉFINI');
      document.getElementById('count-pays-signee').textContent = compPays.length + ' pays';
      renderComparativeHorizontalBarChart('chart-signee-pays', compPays);
      renderComparativeRankedTable('table-signee-pays-body', 'table-signee-pays-head', compPays, 'Pays');

      // --- C. Cédante Comparatif ---
      const compCedantes = aggregateComparative(records26, records25, r => r.cedante || 'NON DÉFINI');
      document.getElementById('count-cedantes-signee').textContent = compCedantes.length + ' groupes';
      renderComparativeHorizontalBarChart('chart-signee-cedante', compCedantes.slice(0, 10));
      renderComparativeRankedTable('table-signee-cedante-body', 'table-signee-cedante-head', compCedantes, 'Cédante (Groupe)');

      // --- D. Courtier Comparatif ---
      const compCourtiers = aggregateComparative(records26, records25, r => r.apporteur || 'NON SPÉCIFIÉ');
      document.getElementById('count-courtiers-signee').textContent = compCourtiers.length + ' apporteurs';
      renderComparativeHorizontalBarChart('chart-signee-courtier', compCourtiers.slice(0, 10));
      renderComparativeRankedTable('table-signee-courtier-body', 'table-signee-courtier-head', compCourtiers, 'Courtier / Apporteur');

      // Top 10 Signées
      renderTop10AffairesSignees();
    }

    function resetStandardTableHead(theadId, labelCol) {
      const thead = document.getElementById(theadId);
      if (!thead) return;
      thead.innerHTML = `
        <tr>
          <th class="py-2 px-3 text-center w-12">Rang</th>
          <th class="py-2 px-3">${labelCol}</th>
          <th class="py-2 px-3 text-right">Prime SCG-Ré</th>
          <th class="py-2 px-3 text-center">Bordereaux</th>
          <th class="py-2 px-3 text-right">Part (%)</th>
        </tr>
      `;
    }

    // ==========================================
    // REUSABLE CHARTS & TABLES
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
            padding: { right: 65, left: 5, top: 5, bottom: 5 }
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
              font: { size: 10, weight: '700', family: 'Plus Jakarta Sans' },
              formatter: (val) => (!val || val === 0 ? '' : formatBarValue(val))
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

    function renderComparativeHorizontalBarChart(canvasId, list) {
      const ctx = document.getElementById(canvasId);
      if (!ctx) return;

      if (charts[canvasId]) {
        charts[canvasId].destroy();
      }

      const topItems = list.slice(0, 10);
      const labels = topItems.map(x => x.name);
      const values26 = topItems.map(x => x.amount2026);
      const values25 = topItems.map(x => x.amount2025);
      const maxVal = Math.max(...values26, ...values25, 0);

      charts[canvasId] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: labels,
          datasets: [
            {
              label: '2026',
              data: values26,
              backgroundColor: '#059669',
              borderRadius: 4,
              borderSkipped: false,
              barPercentage: 0.8,
              categoryPercentage: 0.8
            },
            {
              label: '2025',
              data: values25,
              backgroundColor: '#64748b',
              borderRadius: 4,
              borderSkipped: false,
              barPercentage: 0.8,
              categoryPercentage: 0.8
            }
          ]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          layout: {
            padding: { right: 70, left: 5, top: 5, bottom: 5 }
          },
          plugins: {
            legend: {
              display: true,
              position: 'top',
              align: 'end',
              labels: {
                boxWidth: 12,
                font: { size: 11, weight: '700', family: 'Plus Jakarta Sans' }
              }
            },
            tooltip: {
              callbacks: {
                label: (ctx) => `${ctx.dataset.label} : ${formatFCFA(ctx.raw)}`
              }
            },
            datalabels: {
              anchor: 'end',
              align: 'right',
              offset: 3,
              color: '#334155',
              font: { size: 9, weight: '700', family: 'Plus Jakarta Sans' },
              formatter: (val) => (!val || val === 0 ? '' : formatBarValue(val))
            }
          },
          scales: {
            x: {
              suggestedMax: maxVal > 0 ? maxVal * 1.15 : 100,
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

    function renderComparativeRankedTable(tbodyId, theadId, list, colLabel) {
      const thead = document.getElementById(theadId);
      if (thead) {
        thead.innerHTML = `
          <tr>
            <th class="py-2 px-3 text-center w-10">Rang</th>
            <th class="py-2 px-3">${colLabel}</th>
            <th class="py-2 px-3 text-right text-emerald-800">Prime 2026</th>
            <th class="py-2 px-3 text-right text-slate-600">Prime 2025</th>
            <th class="py-2 px-3 text-right text-slate-700">Évolution</th>
            <th class="py-2 px-3 text-center">Variation</th>
            <th class="py-2 px-3 text-center">Bordereaux (26 vs 25)</th>
          </tr>
        `;
      }

      const tbody = document.getElementById(tbodyId);
      if (!tbody) return;

      if (!list || list.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="py-6 text-center text-slate-400 font-medium">Aucune donnée comparative</td></tr>';
        return;
      }

      tbody.innerHTML = list.map((item, idx) => {
        const rank = idx + 1;
        let rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full text-slate-500 font-bold text-xs">${rank}</span>`;
        if (rank === 1) rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-amber-100 text-amber-800 border border-amber-300 font-extrabold text-xs shadow-xs">1</span>`;
        else if (rank === 2) rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-slate-200 text-slate-700 border border-slate-300 font-extrabold text-xs shadow-xs">2</span>`;
        else if (rank === 3) rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-amber-700/10 text-amber-900 border border-amber-600/30 font-extrabold text-xs shadow-xs">3</span>`;

        const badge = formatVariationBadge(item.amount2026, item.amount2025);
        const deltaSign = item.delta > 0 ? '+' : '';
        const deltaColor = item.delta > 0 ? 'text-emerald-700 font-bold' : (item.delta < 0 ? 'text-rose-700 font-bold' : 'text-slate-500');

        return `
          <tr class="hover:bg-slate-50/80 transition-colors">
            <td class="py-2.5 px-3 font-bold text-center">${rankBadge}</td>
            <td class="py-2.5 px-3 font-semibold text-slate-800">
              <div class="truncate max-w-[180px]" title="${item.name}">${item.name}</div>
            </td>
            <td class="py-2.5 px-3 font-extrabold text-right text-emerald-800 whitespace-nowrap">${formatFCFA(item.amount2026)}</td>
            <td class="py-2.5 px-3 font-medium text-right text-slate-600 whitespace-nowrap">${formatFCFA(item.amount2025)}</td>
            <td class="py-2.5 px-3 text-right whitespace-nowrap text-xs ${deltaColor}">${deltaSign}${formatFCFA(item.delta)}</td>
            <td class="py-2.5 px-3 text-center whitespace-nowrap">${badge}</td>
            <td class="py-2.5 px-3 text-center whitespace-nowrap">
              <span class="px-2 py-0.5 rounded-full text-[11px] font-semibold bg-slate-100 text-slate-700 border border-slate-200/60">${item.count2026} vs ${item.count2025}</span>
            </td>
          </tr>
        `;
      }).join('');
    }

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
        let rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full text-slate-500 font-bold text-xs">${rank}</span>`;
        if (rank === 1) rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-amber-100 text-amber-800 border border-amber-300 font-extrabold text-xs shadow-xs">1</span>`;
        else if (rank === 2) rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-slate-200 text-slate-700 border border-slate-300 font-extrabold text-xs shadow-xs">2</span>`;
        else if (rank === 3) rankBadge = `<span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-amber-700/10 text-amber-900 border border-amber-600/30 font-extrabold text-xs shadow-xs">3</span>`;

        return `
          <tr class="hover:bg-emerald-50/40 transition-colors">
            <td class="py-2.5 px-3 font-bold text-center">${rankBadge}</td>
            <td class="py-2.5 px-3 font-extrabold text-slate-900">
              <div class="truncate max-w-[220px]" title="${r.affaire}">${r.affaire}</div>
            </td>
            <td class="py-2.5 px-3 font-semibold text-slate-700">
              <div class="truncate max-w-[160px]" title="${r.cedante}">${r.cedante}</div>
            </td>
            <td class="py-2.5 px-3 text-slate-600">
              <div class="truncate max-w-[140px]" title="${r.apporteur}">${r.apporteur}</div>
            </td>
            <td class="py-2.5 px-3 text-slate-600 whitespace-nowrap">${r.branche || '-'}</td>
            <td class="py-2.5 px-3 text-right font-medium text-slate-600 whitespace-nowrap">${formatFCFA(r.capitaux100)}</td>
            <td class="py-2.5 px-3 text-right font-medium text-slate-600 whitespace-nowrap">${formatFCFA(r.primes100)}</td>
            <td class="py-2.5 px-3 text-center whitespace-nowrap">
              <span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-slate-100 text-slate-700">${formatPercent(r.tauxCession)}</span>
            </td>
            <td class="py-2.5 px-3 text-right font-black text-emerald-700 whitespace-nowrap text-sm bg-emerald-50/50">${formatFCFA(r.primesSCG)}</td>
          </tr>
        `;
      }).join('');
    }

    // ==========================================
    // 4. VUE AFFAIRES SANS SUITE (EXERCICE 2026)
    // ==========================================
    function renderVueSansSuite() {
      const ssRecords = filteredRecords.filter(r => r.isSansSuite === true);
      const totalPrimes100 = ssRecords.reduce((acc, r) => acc + (r.primes100 || 0), 0);
      const totalCapitaux100 = ssRecords.reduce((acc, r) => acc + (r.capitaux100 || 0), 0);
      const total2026Affaires = filteredRecords.length;
      const total2026Primes100 = filteredRecords.reduce((acc, r) => acc + (r.primes100 || 0), 0);
      const tauxDeperditionNombre = total2026Affaires > 0 ? ((ssRecords.length / total2026Affaires) * 100) : 0;
      const tauxDeperditionPrime = total2026Primes100 > 0 ? ((totalPrimes100 / total2026Primes100) * 100) : 0;
      const moyPrime = ssRecords.length > 0 ? (totalPrimes100 / ssRecords.length) : 0;

      // Banner badges
      const bCount = document.getElementById('badge-sans-suite-count');
      if (bCount) bCount.textContent = `${ssRecords.length} affaires sans suite`;
      const bTotal = document.getElementById('badge-sans-suite-total');
      if (bTotal) bTotal.textContent = `Primes 100% : ${formatFCFA(totalPrimes100)}`;

      // 4 KPIs
      const elCount = document.getElementById('kpi-ss-count');
      if (elCount) elCount.textContent = ssRecords.length;
      const elTaux = document.getElementById('kpi-ss-taux-deperdition');
      if (elTaux) elTaux.textContent = `${tauxDeperditionNombre.toFixed(1)}%`;

      const elPrimes = document.getElementById('kpi-ss-primes-100');
      if (elPrimes) elPrimes.textContent = formatFCFA(totalPrimes100);
      const elTauxPrime = document.getElementById('kpi-ss-taux-prime');
      if (elTauxPrime) elTauxPrime.textContent = `${tauxDeperditionPrime.toFixed(1)}%`;

      const elCapitaux = document.getElementById('kpi-ss-capitaux-100');
      if (elCapitaux) elCapitaux.textContent = formatFCFA(totalCapitaux100);

      const elMoy = document.getElementById('kpi-ss-moyenne-prime');
      if (elMoy) elMoy.textContent = formatFCFA(moyPrime);

      // --- A. Branche Sans Suite ---
      const sortedBranches = aggregateBySum(ssRecords, r => r.branche || 'AUTRES', r => r.primes100 || 0);
      document.getElementById('count-branches-ss').textContent = `${sortedBranches.length} branches`;
      renderHorizontalBarChart('chart-ss-branche', {
        labels: sortedBranches.map(x => x.name),
        values: sortedBranches.map(x => x.amount),
        color: '#e11d48',
        datasetLabel: 'Primes 100% sans suite'
      });
      renderRankedTable('table-ss-branche-body', sortedBranches, totalPrimes100, 'affaires', 'bg-rose-600');

      // --- B. Pays Sans Suite ---
      const sortedPays = aggregateBySum(ssRecords, r => r.pays || 'NON DÉFINI', r => r.primes100 || 0);
      document.getElementById('count-pays-ss').textContent = `${sortedPays.length} pays`;
      renderHorizontalBarChart('chart-ss-pays', {
        labels: sortedPays.map(x => x.name),
        values: sortedPays.map(x => x.amount),
        color: '#f43f5e',
        datasetLabel: 'Primes 100% sans suite'
      });
      renderRankedTable('table-ss-pays-body', sortedPays, totalPrimes100, 'affaires', 'bg-rose-500');

      // --- C. Cédante Sans Suite ---
      const sortedCedantes = aggregateBySum(ssRecords, r => r.cedante || 'NON DÉFINI', r => r.primes100 || 0);
      document.getElementById('count-cedantes-ss').textContent = `${sortedCedantes.length} groupes`;
      renderHorizontalBarChart('chart-ss-cedante', {
        labels: sortedCedantes.slice(0, 10).map(x => x.name),
        values: sortedCedantes.slice(0, 10).map(x => x.amount),
        color: '#d97706',
        datasetLabel: 'Primes 100% sans suite'
      });
      renderRankedTable('table-ss-cedante-body', sortedCedantes, totalPrimes100, 'affaires', 'bg-amber-600');

      // --- D. Courtier Sans Suite ---
      const sortedCourtiers = aggregateBySum(ssRecords, r => r.apporteur || 'NON SPÉCIFIÉ', r => r.primes100 || 0);
      document.getElementById('count-courtiers-ss').textContent = `${sortedCourtiers.length} apporteurs`;
      renderHorizontalBarChart('chart-ss-courtier', {
        labels: sortedCourtiers.slice(0, 10).map(x => x.name),
        values: sortedCourtiers.slice(0, 10).map(x => x.amount),
        color: '#ea580c',
        datasetLabel: 'Primes 100% sans suite'
      });
      renderRankedTable('table-ss-courtier-body', sortedCourtiers, totalPrimes100, 'affaires', 'bg-orange-600');

      // Detailed Table
      renderDetailedSansSuiteTable(ssRecords);
    }

    function setupSansSuiteControls() {
      const inputSearch = document.getElementById('input-search-ss');
      if (inputSearch) {
        inputSearch.addEventListener('input', (e) => {
          searchSS = e.target.value.toLowerCase().trim();
          const ssRecords = filteredRecords.filter(r => r.isSansSuite === true);
          renderDetailedSansSuiteTable(ssRecords);
        });
      }

      const btnExport = document.getElementById('btn-export-ss');
      if (btnExport) {
        btnExport.addEventListener('click', () => {
          const ssRecords = filteredRecords.filter(r => r.isSansSuite === true);
          exportSansSuiteExcel(ssRecords);
        });
      }
    }

    function renderDetailedSansSuiteTable(ssRecords) {
      const tbody = document.getElementById('table-sans-suite-detailed-body');
      if (!tbody) return;

      let list = [...ssRecords];
      if (searchSS) {
        list = list.filter(r => 
          (r.affaire && r.affaire.toLowerCase().includes(searchSS)) ||
          (r.cedante && r.cedante.toLowerCase().includes(searchSS)) ||
          (r.apporteur && r.apporteur.toLowerCase().includes(searchSS)) ||
          (r.branche && r.branche.toLowerCase().includes(searchSS)) ||
          (r.pays && r.pays.toLowerCase().includes(searchSS)) ||
          (r.commentaires && r.commentaires.toLowerCase().includes(searchSS))
        );
      }

      // Tri par Prime 100% décroissante
      list.sort((a, b) => (b.primes100 || 0) - (a.primes100 || 0));

      if (list.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="py-6 text-center text-slate-400 font-medium">Aucune affaire classée sans suite trouvée</td></tr>';
        return;
      }

      tbody.innerHTML = list.map(r => {
        return `
          <tr class="hover:bg-rose-50/40 transition-colors">
            <td class="py-2.5 px-3 font-semibold text-slate-500 text-center">${r.id}</td>
            <td class="py-2.5 px-3 text-slate-600 whitespace-nowrap">${r.dateSouscription || '-'}</td>
            <td class="py-2.5 px-3 font-bold text-slate-900">
              <div class="truncate max-w-[200px]" title="${r.affaire}">${r.affaire}</div>
            </td>
            <td class="py-2.5 px-3 font-medium text-slate-700">
              <div class="truncate max-w-[160px]" title="${r.cedante}">${r.cedante}</div>
            </td>
            <td class="py-2.5 px-3 text-slate-600">
              <div class="truncate max-w-[140px]" title="${r.apporteur}">${r.apporteur}</div>
            </td>
            <td class="py-2.5 px-3 text-slate-600 whitespace-nowrap">${r.branche || '-'}</td>
            <td class="py-2.5 px-3 text-slate-600 whitespace-nowrap">${r.pays || '-'}</td>
            <td class="py-2.5 px-3 text-right font-medium text-slate-700 whitespace-nowrap">${formatFCFA(r.capitaux100)}</td>
            <td class="py-2.5 px-3 text-right font-bold text-amber-700 whitespace-nowrap bg-amber-50/40">${formatFCFA(r.primes100)}</td>
            <td class="py-2.5 px-3">
              <span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-100 text-rose-800 border border-rose-200">
                ${r.commentaires || 'CLASSEE SANS SUITE'}
              </span>
            </td>
          </tr>
        `;
      }).join('');
    }

    function exportSansSuiteExcel(ssRecords) {
      if (!ssRecords || ssRecords.length === 0) {
        alert("Aucune donnée à exporter.");
        return;
      }
      const rows = ssRecords.map(r => ({
        "N° ID": r.id,
        "Exercice": r.exercice || 2026,
        "Date Souscription": r.dateSouscription,
        "Affaire": r.affaire,
        "Cédante": r.cedante,
        "Courtier": r.apporteur,
        "Branche": r.branche,
        "Pays": r.pays,
        "Capitaux 100%": r.capitaux100,
        "Primes 100%": r.primes100,
        "Motif / Statut": r.commentaires || 'CLASSEE SANS SUITE'
      }));
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Affaires Sans Suite");
      XLSX.writeFile(wb, "SCG_Re_Affaires_Sans_Suite_2026.xlsx");
      showToast("Fichier Excel des affaires sans suite exporté !");
    }

    // ==========================================
    // 5. DATA EXPLORER & EXPORT (ONGLET 4)
    // ==========================================
    function setupTableControls() {
      const searchInput = document.getElementById('table-search');
      const pageSizeSelect = document.getElementById('table-page-size');
      const btnExport = document.getElementById('btn-export-excel');
      const btnPrev = document.getElementById('btn-page-prev');
      const btnNext = document.getElementById('btn-page-next');

      if (searchInput) {
        searchInput.addEventListener('input', (e) => {
          searchTerm = e.target.value.toLowerCase().trim();
          currentPage = 1;
          renderDataTable();
        });
      }

      if (pageSizeSelect) {
        pageSizeSelect.addEventListener('change', (e) => {
          pageSize = parseInt(e.target.value, 10);
          currentPage = 1;
          renderDataTable();
        });
      }

      if (btnExport) {
        btnExport.addEventListener('click', exportCurrentTable);
      }

      if (btnPrev) {
        btnPrev.addEventListener('click', () => changePage(-1));
      }

      if (btnNext) {
        btnNext.addEventListener('click', () => changePage(1));
      }

      // Column headers click-to-sort
      document.querySelectorAll('#main-data-table thead th[data-sort]').forEach(th => {
        th.addEventListener('click', () => {
          const col = th.getAttribute('data-sort');
          if (sortColumn === col) {
            sortDirection = -sortDirection;
          } else {
            sortColumn = col;
            sortDirection = 1;
          }
          renderDataTable();
        });
      });
    }

    function renderDataTable() {
      // Choix du dataset selon l'exercice sélectionné
      let base = (explorerExercice === '2025') 
        ? records2025.filter(r => !selectedBranche || r.branche === selectedBranche)
        : filteredRecords;

      let data = [...base];

      // Text search
      if (searchTerm) {
        data = data.filter(r => {
          return (r.affaire && r.affaire.toLowerCase().includes(searchTerm)) ||
                 (r.cedante && r.cedante.toLowerCase().includes(searchTerm)) ||
                 (r.apporteur && r.apporteur.toLowerCase().includes(searchTerm)) ||
                 (r.pays && r.pays.toLowerCase().includes(searchTerm)) ||
                 (r.branche && r.branche.toLowerCase().includes(searchTerm)) ||
                 (r.etatBordereau && r.etatBordereau.toLowerCase().includes(searchTerm)) ||
                 (r.commentaires && r.commentaires.toLowerCase().includes(searchTerm)) ||
                 (r.id && r.id.toString().includes(searchTerm));
        });
      }

      document.getElementById('table-records-found').textContent = data.length + ' polices (' + explorerExercice + ')';

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
          let statusBadge = '';
          if (r.isSansSuite) {
            statusBadge = `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-100 text-rose-800 border border-rose-200">SANS SUITE</span>`;
          } else if (r.etatBordereau === 'REÇU') {
            statusBadge = `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">SIGNÉE</span>`;
          } else {
            statusBadge = `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-800 border border-amber-200">EN ATTENTE</span>`;
          }

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
                ${statusBadge}
              </td>
            </tr>
          `;
        }).join('');
      }

      // Update Pagination UI
      const endIdx = Math.min(startIdx + pageSize, data.length);
      document.getElementById('table-pagination-info').textContent = 
        data.length > 0 ? `Affichage de ${startIdx + 1} à ${endIdx} sur ${data.length} affaires (${explorerExercice})` : '0 affaire';
      document.getElementById('table-current-page').textContent = `Page ${currentPage} / ${totalPages}`;
      document.getElementById('btn-page-prev').disabled = (currentPage <= 1);
      document.getElementById('btn-page-next').disabled = (currentPage >= totalPages);
    }

    function changePage(delta) {
      currentPage += delta;
      renderDataTable();
    }

    function exportCurrentTable() {
      const base = (explorerExercice === '2025') 
        ? records2025.filter(r => !selectedBranche || r.branche === selectedBranche)
        : filteredRecords;

      let data = [...base];
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

      if (data.length === 0) {
        alert('Aucune donnée à exporter.');
        return;
      }

      const rows = data.map(r => ({
        "ID": r.id,
        "Exercice": r.exercice || (explorerExercice === '2025' ? 2025 : 2026),
        "Date Souscription": r.dateSouscription,
        "Date Réception Bordereau": r.dateReception,
        "Délai (jours)": r.delaiReceptionJours,
        "Affaire": r.affaire,
        "Type Affaire": r.typeAffaire,
        "Cédante": r.cedante,
        "Courtier": r.apporteur,
        "Pays": r.pays,
        "Branche": r.branche,
        "Capitaux 100%": r.capitaux100,
        "Primes 100%": r.primes100,
        "Capitaux SCG-Ré": r.capitauxSCG,
        "Primes SCG-Ré": r.primesSCG,
        "Commission": r.commission,
        "Primes Nettes": r.primesNettes,
        "État Bordereau": r.etatBordereau,
        "Classée Sans Suite": r.isSansSuite ? "OUI" : "NON",
        "Commentaires": r.commentaires
      }));

      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, `Production_${explorerExercice}`);
      XLSX.writeFile(wb, `SCG_Re_Production_${explorerExercice}_Export.xlsx`);
      showToast(`Exportation Excel (${explorerExercice}) réussie !`);
    }

    // ==========================================
    // PERSISTENCE CONTROLS & DOWNLOAD HTML
    // ==========================================
    function setupStorageControls() {
      const btnReset = document.getElementById('btn-reset-storage');
      if (btnReset) {
        btnReset.addEventListener('click', async () => {
          if (confirm("Voulez-vous rétablir les données d'origine pré-intégrées ?")) {
            await clearDatasetFromStorage();
            allRecords = [...INITIAL_RECORDS_2026];
            records2025 = [...INITIAL_RECORDS_2025];
            filteredRecords = [...allRecords];
            selectedBranche = '';
            document.getElementById('filter-branche').value = '';
            populateFilterOptions();
            renderAll();
            showToast("Données d'origine rétablies.");
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
        let fullHtml = document.documentElement.outerHTML;
        const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'index.html';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast('Fichier index.html complet téléchargé !');
      } catch (err) {
        alert("Erreur lors du téléchargement : " + err.message);
      }
    }

    // ==========================================
    // IN-BROWSER EXCEL IMPORT & CLEANING
    // ==========================================
    function cleanBranchJS(b) {
      if (!b) return 'AUTRES';
      const up = String(b).replace(/\\u00a0/g, ' ').trim().toUpperCase();
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

    function cleanCountryJS(p) {
      if (!p) return 'NON DÉFINI';
      const up = String(p).replace(/\\u00a0/g, ' ').trim().toUpperCase();
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
      const s = String(app).trim();
      const up = s.toUpperCase();
      if (up.includes('GLUBUS')) return 'GLOBUS RE';
      if (up.includes('PREVEX')) return 'PREVEX RE';
      if (up.includes('ATLAS REINSURANCE') || up === 'ARC') return 'ATLAS REINSURANCE CONSULTANTS';
      if (up === 'SBRE' || up === 'SB RE') return 'SB RE';
      if (up.includes('KLAPTON')) return 'KLAPTON AFRICA';
      if (up === 'SUNU GROUP' || up.includes('SUNU REASSURANCE') || up.includes('SUNU RE')) return 'SUNU RE';
      return s;
    }

    function cleanCedanteJS(ced) {
      if (!ced || String(ced).trim() === '') return 'NON SPÉCIFIÉ';
      const s = String(ced).trim();
      const up = s.toUpperCase();
      if (up === 'TBA') return 'TBA';
      if (up.includes(',') || (up.includes('&') && !up.includes('ALLIANZ') && !up.includes('SANLAM'))) return 'CO-ASSURANCE / APÉRITEUR';
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
      if (up.includes('UAB')) return 'UAB';
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

    function isSansSuiteJS(comment) {
      if (!comment) return false;
      const norm = String(comment).normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase();
      return norm.includes('sans suite') || norm.includes('classe sans suite');
    }

    function setupUploadModal() {
      const modal = document.getElementById('upload-modal');
      const btnOpen = document.getElementById('btn-open-upload');
      const btnClose = document.getElementById('btn-close-upload');
      const btnCancel = document.getElementById('btn-cancel-upload');
      const dropZone = document.getElementById('drop-zone');
      const fileInput = document.getElementById('file-input');

      if (btnOpen) btnOpen.addEventListener('click', () => modal.classList.remove('hidden'));
      if (btnClose) btnClose.addEventListener('click', () => modal.classList.add('hidden'));
      if (btnCancel) btnCancel.addEventListener('click', () => modal.classList.add('hidden'));

      if (dropZone && fileInput) {
        dropZone.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => {
          if (e.target.files.length > 0) processExcelFile(e.target.files[0]);
        });
        dropZone.addEventListener('dragover', (e) => {
          e.preventDefault();
          dropZone.classList.add('border-brand-600', 'bg-brand-100/50');
        });
        dropZone.addEventListener('dragleave', () => {
          dropZone.classList.remove('border-brand-600', 'bg-brand-100/50');
        });
        dropZone.addEventListener('drop', (e) => {
          e.preventDefault();
          dropZone.classList.remove('border-brand-600', 'bg-brand-100/50');
          if (e.dataTransfer.files.length > 0) processExcelFile(e.dataTransfer.files[0]);
        });
      }
    }

    function processExcelFile(file) {
      const reader = new FileReader();
      reader.onload = async function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          
          let targetSheetName = workbook.SheetNames[0];
          for (const name of workbook.SheetNames) {
            const upN = name.toUpperCase();
            if (upN.includes('BASE') || upN.includes('PRODUCTION')) {
              targetSheetName = name;
              break;
            }
          }
          const sheet = workbook.Sheets[targetSheetName];
          const rawMatrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
          if (!rawMatrix || rawMatrix.length === 0) {
            alert('Le fichier Excel semble vide.');
            return;
          }

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

          if (headerIdx === -1) headerIdx = 0;
          const headers = rawMatrix[headerIdx].map(h => String(h || '').trim());
          
          const findColIdx = (...candidates) => {
            const normC = candidates.map(c => c.toLowerCase().replace(/[^a-z0-9]/g, ''));
            // Exact pass
            const exactIdx = headers.findIndex(h => {
              const normH = h.toLowerCase().replace(/[^a-z0-9]/g, '');
              return normC.some(c => normH === c);
            });
            if (exactIdx !== -1) return exactIdx;
            // Inclusion pass
            return headers.findIndex(h => {
              const normH = h.toLowerCase().replace(/[^a-z0-9]/g, '');
              return normC.some(c => normH.includes(c) || c.includes(normH));
            });
          };

          const colDateSouscr = findColIdx('date de souscription', 'datesouscription');
          const colDateRecept = findColIdx('date de reception du bordereau', 'datereception', 'date de reception');
          const colAffaire = findColIdx('affaires', 'affaire');
          const colTypeAffaire = findColIdx("type d'affaire", 'typeaffaire');
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
          const colComment = findColIdx('commentaires', 'commentaire');

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

          const isFile2025 = file.name.includes('2025');
          const fileYear = isFile2025 ? 2025 : 2026;

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

            const commentVal = colComment !== -1 ? String(row[colComment] || '').trim() : '';
            const ssFlag = isSansSuiteJS(commentVal);

            newRecords.push({
              id: counter++,
              exercice: fileYear,
              isSansSuite: ssFlag,
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
              etatBordereau: dRecept ? 'REÇU' : 'EN ATTENTE',
              commentaires: commentVal
            });
          }

          if (newRecords.length === 0) {
            alert("Aucune ligne de production valide n'a pu être extraite.");
            return;
          }

          if (isFile2025) {
            records2025 = newRecords;
            showToast(`Fichier 2025 importé : ${records2025.length} affaires enregistrées !`);
          } else {
            allRecords = newRecords;
            filteredRecords = [...allRecords];
            populateFilterOptions();
            showToast(`Fichier 2026 importé : ${allRecords.length} affaires enregistrées !`);
          }

          await saveDatasetToStorage({
            records2026: allRecords,
            records2025: records2025,
            fileName: file.name,
            importDate: new Date().toLocaleString('fr-FR')
          });

          renderAll();
          document.getElementById('upload-modal').classList.add('hidden');
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

final_html = html_template.replace('__INITIAL_RECORDS_2026__', records_2026_json).replace('__INITIAL_RECORDS_2025__', records_2025_json).replace('__INITIAL_CLEANING_LOG__', log_json).replace('__LOGO_SRC__', logo_src).replace('__BUILD_TIMESTAMP__', datetime.now().isoformat())

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(final_html)

print("index.html generated successfully with all new features (Affaires Sans Suite & Comparatif 2025 vs 2026)!")
