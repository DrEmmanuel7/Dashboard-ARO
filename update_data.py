import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import json
import re
import os
import glob
import unicodedata

def normalize_text(text):
    if not text:
        return ''
    text = unicodedata.normalize('NFD', str(text))
    return ''.join(c for c in text if unicodedata.category(c) != 'Mn').lower()

def is_sans_suite(comment):
    norm = normalize_text(comment)
    return 'sans suite' in norm or 'classe sans suite' in norm

def excel_date_to_str(val):
    if not val:
        return None
    val_str = str(val).strip()
    if not val_str:
        return None
    try:
        f = float(val_str)
        dt = datetime(1899, 12, 30) + timedelta(days=f)
        return dt.strftime('%Y-%m-%d')
    except:
        return val_str

def clean_branche(b):
    if not b:
        return 'AUTRES'
    up = b.replace('\u00a0', ' ').strip().upper()
    if 'DECENNALE' in up or 'DÉCENNALE' in up:
        return 'RC DÉCENNALE'
    if 'INCENDIE' in up:
        return 'INCENDIE'
    if 'CIVILE' in up:
        return 'RESPONSABILITÉ CIVILE'
    if 'TECHNIQUE' in up:
        return 'RISQUES TECHNIQUES'
    if 'TRANSPORT' in up:
        return 'TRANSPORTS'
    if 'DIVERS' in up:
        return 'RISQUES DIVERS'
    if 'ACCIDENT' in up:
        return 'INDIVIDUELLE ACCIDENTS'
    if 'AUTO' in up:
        return 'AUTOMOBILE'
    if 'CAUTION' in up:
        return 'CAUTION'
    if 'ENERGIE' in up or 'ÉNERGIE' in up:
        return 'ÉNERGIE'
    if 'PVT' in up or 'TERRORISME' in up or 'VIOLENCE' in up:
        return 'PVT'
    return up

def clean_pays(p):
    if not p:
        return 'NON DÉFINI'
    up = p.replace('\u00a0', ' ').strip().upper()
    if 'COTE' in up or 'CÔTE' in up:
        if 'SENEGAL' in up or 'SÉNÉGAL' in up:
            return "CÔTE D'IVOIRE / SÉNÉGAL"
        return "CÔTE D'IVOIRE"
    if 'BENIN' in up or 'BÉNIN' in up:
        return 'BÉNIN'
    if 'BURKINA' in up:
        return 'BURKINA FASO'
    if 'SENEGAL' in up or 'SÉNÉGAL' in up:
        return 'SÉNÉGAL'
    if 'MALI' in up:
        return 'MALI'
    if 'TOGO' in up:
        return 'TOGO'
    if 'NIGER' in up:
        return 'NIGER'
    if 'GABON' in up:
        return 'GABON'
    if 'CAMEROUN' in up:
        return 'CAMEROUN'
    if 'GUINEE' in up or 'GUINÉE' in up:
        if 'EQUATORIAL' in up or 'ÉQUATORIALE' in up:
            return 'GUINÉE ÉQUATORIALE'
        return 'GUINÉE CONAKRY'
    if 'CONGO' in up:
        return 'CONGO'
    if 'TUNISIE' in up:
        return 'TUNISIE'
    if 'MAROC' in up:
        return 'MAROC'
    if 'ALGERIE' in up or 'ALGÉRIE' in up:
        return 'ALGÉRIE'
    if 'MAURITANIE' in up:
        return 'MAURITANIE'
    if 'TCHAD' in up:
        return 'TCHAD'
    if 'CENTRAFRIQUE' in up or 'RCA' in up:
        return 'CENTRAFRIQUE'
    if 'FRANCE' in up:
        return 'FRANCE'
    if 'CAP VERT' in up or 'CABO VERDE' in up:
        return 'CAP-VERT'
    if 'GHANA' in up:
        return 'GHANA'
    return up

def clean_apporteur(app):
    if not app:
        return 'NON SPÉCIFIÉ'
    s = str(app).strip()
    up = s.upper()
    if 'GLUBUS' in up:
        return 'GLOBUS RE'
    if 'PREVEX' in up:
        return 'PREVEX RE'
    if 'ATLAS REINSURANCE' in up or up == 'ARC':
        return 'ATLAS REINSURANCE CONSULTANTS'
    if up in ['SBRE', 'SB RE']:
        return 'SB RE'
    if 'KLAPTON' in up:
        return 'KLAPTON AFRICA'
    if up == 'SUNU GROUP' or 'SUNU REASSURANCE' in up or 'SUNU RE' in up:
        return 'SUNU RE'
    return s

def clean_cedante(ced):
    if not ced or str(ced).strip() == '':
        return 'NON SPÉCIFIÉ'
    s = str(ced).strip()
    up = s.upper()
    if up == 'TBA':
        return 'TBA'
    if ',' in up or ('&' in up and not 'ALLIANZ' in up and not 'SANLAM' in up):
        return 'CO-ASSURANCE / APÉRITEUR'
    if 'SANLAMALLIANZ & AMSA' in up or 'SANLAM ALLIANZ,AMSA' in up:
        return 'CO-ASSURANCE / APÉRITEUR'
    if 'SUNU' in up:
        return 'SUNU ASSURANCES'
    if 'GNA' in up:
        return 'GNA'
    if 'NSIA' in up:
        return 'NSIA'
    if 'CORIS' in up:
        return 'CORIS ASSURANCES'
    if 'SANLAM' in up or 'ALLIANZ' in up:
        return 'SANLAM ALLIANZ'
    if 'ATLANTIQUE' in up:
        return 'ATLANTIQUE ASSURANCES'
    if 'ACTIVA' in up:
        return 'ACTIVA'
    if 'AFG' in up:
        return 'AFG'
    if 'AMSA' in up or 'ASMA' in up:
        return 'AMSA'
    if 'VISTA' in up:
        return 'VISTA ASSURANCES'
    if 'RAYNAL' in up:
        return 'RAYNAL ASSURANCES'
    if 'FINAFRICA' in up:
        return 'FINAFRICA ASSURANCES'
    if 'LANALA' in up:
        return 'LANALA ASSURANCES'
    if 'BIAT' in up:
        return 'ASSURANCES BIAT'
    if 'ASCOMA' in up:
        return 'ASCOMA'
    if 'CARTE' in up:
        return 'CARTE ASSURANCES'
    if 'ASTREE' in up or 'ASTRÉE' in up:
        return 'ASTRÉE ASSURANCES'
    if 'ATLANTA' in up:
        return 'ATLANTA'
    if 'SONAR' in up:
        return 'SONAR'
    if 'FASO RE' in up:
        return 'FASO RE'
    if 'LEADWAY' in up:
        return 'LEADWAY ASSURANCES'
    if 'ASKIA' in up:
        return 'ASKIA ASSURANCES'
    if 'ENERALE' in up or 'GÉNÉRALE' in up:
        return 'GÉNÉRALE DES ASSURANCES'
    if 'L\'AFRICAINE' in up:
        return 'L\'AFRICAINE DES ASSURANCES'
    if 'UNION DES ASSURANCES DU BURKINA' in up or 'UAB' in up:
        return 'UAB'
    if 'ZITOUNA' in up:
        return 'ZITOUNA TAKAFUL'
    if 'BH' in up:
        return 'BH ASSURANCES'
    if 'GAT' in up:
        return 'GAT ASSURANCES'
    if 'PRUDENTIAL' in up:
        return 'PRUDENTIAL BENEFICIAL'
    if 'IMPAR' in up:
        return 'IMPAR'
    if 'COMAR' in up:
        return 'COMAR'
    if 'SERENITY' in up:
        return 'SERENITY SA'
    if 'FIDELIA' in up:
        return 'FIDELIA ASSURANCES'
    if 'PREVOYANCE' in up:
        return 'PRÉVOYANCE ASSURANCES'
    if 'GEPETROL' in up:
        return 'GEPETROL SEGUROS'
    return s

def to_float(val):
    if not val:
        return 0.0
    try:
        return float(str(val).replace(' ', '').replace(',', '.'))
    except:
        return 0.0

def parse_excel_file(filepath, exercice_year, start_id=1):
    print(f"Traitement de l'exercice {exercice_year} : {filepath}")
    with zipfile.ZipFile(filepath, 'r') as z:
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ss_xml = z.read('xl/sharedStrings.xml')
            ss_root = ET.fromstring(ss_xml)
            for si in ss_root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                t = si.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                shared_strings.append(t.text if t is not None else '')

        workbook_xml = z.read('xl/workbook.xml')
        wb_root = ET.fromstring(workbook_xml)
        sheet_elems = list(wb_root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'))

        target_sheet_file = 'xl/worksheets/sheet1.xml'
        for s in sheet_elems:
            name = s.attrib.get('name', '').upper()
            if 'BASE' in name or 'PRODUCTION' in name:
                r_id = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if r_id and r_id.startswith('rId'):
                    num = r_id.replace('rId', '')
                    target_sheet_file = f'xl/worksheets/sheet{num}.xml'
                break

        sheet_xml = z.read(target_sheet_file)
        s_root = ET.fromstring(sheet_xml)
        rows = list(s_root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'))

    header_row_num = None
    header_col_map = {}
    for r in rows:
        row_num = int(r.attrib.get('r'))
        vals = {}
        for c in r.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            ref = ''.join([ch for ch in c.attrib.get('r') if ch.isalpha()])
            t_attr = c.attrib.get('t')
            v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            val = v.text if v is not None else ''
            if t_attr == 's' and val.isdigit():
                val = shared_strings[int(val)]
            vals[ref] = str(val).strip()

        line_vals = [x.upper() for x in vals.values() if x]
        if any('DATE DE SOUSCRIPTION' in x for x in line_vals):
            header_row_num = row_num
            header_col_map = vals
            break

    if not header_row_num:
        header_row_num = 1
        for c in rows[0].iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            ref = ''.join([ch for ch in c.attrib.get('r') if ch.isalpha()])
            t_attr = c.attrib.get('t')
            v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            val = v.text if v is not None else ''
            if t_attr == 's' and val.isdigit():
                val = shared_strings[int(val)]
            header_col_map[ref] = str(val).strip()

    # Map normalized header names to column letters
    norm_headers = {}
    for col, h in header_col_map.items():
        norm = normalize_text(h)
        norm_headers[norm] = col

    def get_val(row_dict, *candidates):
        # 1. Exact match pass
        for cand in candidates:
            norm_cand = normalize_text(cand)
            for norm_h, col in norm_headers.items():
                if norm_cand == norm_h:
                    if col in row_dict:
                        return row_dict[col]
        # 2. Inclusion match fallback
        for cand in candidates:
            norm_cand = normalize_text(cand)
            for norm_h, col in norm_headers.items():
                if norm_cand in norm_h:
                    if col in row_dict:
                        return row_dict[col]
        return ''

    records = []
    counter = start_id
    for r in rows:
        row_num = int(r.attrib.get('r'))
        if row_num <= header_row_num:
            continue

        row_dict = {}
        for c in r.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            ref = ''.join([ch for ch in c.attrib.get('r') if ch.isalpha()])
            t_attr = c.attrib.get('t')
            v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            val = v.text if v is not None else ''
            if t_attr == 's' and val.isdigit():
                val = shared_strings[int(val)]
            row_dict[ref] = str(val).strip()

        affaire = get_val(row_dict, 'affaires', 'affaire').strip()
        id_val = get_val(row_dict, 'n° id', 'id', 'n id').strip()
        if not affaire and not id_val:
            continue

        date_souscr = excel_date_to_str(get_val(row_dict, 'date de souscription', 'datesouscription'))
        date_recept = excel_date_to_str(get_val(row_dict, 'date de reception du bordereau', 'datereception', 'date de reception'))

        delai = None
        if date_souscr and date_recept:
            try:
                d1 = datetime.strptime(date_souscr, '%Y-%m-%d')
                d2 = datetime.strptime(date_recept, '%Y-%m-%d')
                delai = (d2 - d1).days
            except:
                delai = None

        cap100 = to_float(get_val(row_dict, 'capitaux 100 %', 'capitaux 100%'))
        prim100 = to_float(get_val(row_dict, 'primes 100 %', 'primes 100%'))
        capSCG = to_float(get_val(row_dict, 'capitaux scg-re', 'capitaux scg'))
        primSCG = to_float(get_val(row_dict, 'primes scg-re', 'primes scg'))
        comm = to_float(get_val(row_dict, 'commission'))
        taxe = to_float(get_val(row_dict, 'taxe'))
        primNet = to_float(get_val(row_dict, 'primes nettes', 'primes nettes de reassurance'))
        if primNet == 0.0 and primSCG > 0:
            primNet = primSCG - comm

        comment = get_val(row_dict, 'commentaires', 'commentaire').strip()
        ss_flag = is_sans_suite(comment)

        records.append({
            'id': counter,
            'exercice': exercice_year,
            'isSansSuite': ss_flag,
            'dateSouscription': date_souscr,
            'dateReception': date_recept,
            'delaiReceptionJours': delai,
            'typeCession': get_val(row_dict, 'type de cession') or 'CESSION COMMERCIALE FAC',
            'affaire': affaire,
            'typeAffaire': get_val(row_dict, "type d'affaire", 'typeaffaire'),
            'operations': get_val(row_dict, 'operations'),
            'cedante': clean_cedante(get_val(row_dict, 'cedante', 'cedantes')),
            'apporteur': clean_apporteur(get_val(row_dict, 'apporteur', 'apporteurs')),
            'typeApporteur': get_val(row_dict, 'type apporteur'),
            'pays': clean_pays(get_val(row_dict, 'pays (lieu du risque)', 'pays', 'lieu du risque')),
            'bureau': get_val(row_dict, 'bureau'),
            'zoneAssurance': get_val(row_dict, 'zone assurance'),
            'zoneGeo': get_val(row_dict, 'zone geographique'),
            'branche': clean_branche(get_val(row_dict, 'branches', 'branche')),
            'garanties': get_val(row_dict, 'garanties'),
            'typeAssurance': get_val(row_dict, "types d'assurance"),
            'dateEffet': excel_date_to_str(get_val(row_dict, "date d'effet")),
            'dateEcheance': excel_date_to_str(get_val(row_dict, "date d'echeance")),
            'devise': get_val(row_dict, 'devises') or 'XOF',
            'tauxChange': to_float(get_val(row_dict, 'taux de changes')) or 1.0,
            'tauxCession': to_float(get_val(row_dict, 'taux  de cession', 'taux de cession')),
            'tauxCommission': to_float(get_val(row_dict, 'taux de commission')),
            'capitaux100': cap100,
            'primes100': prim100,
            'capitauxSCG': capSCG,
            'primesSCG': primSCG,
            'commission': comm,
            'taxe': taxe,
            'primesNettes': primNet,
            'etatBordereau': 'REÇU' if date_recept else 'EN ATTENTE',
            'commentaires': comment
        })
        counter += 1

    return records

# ==========================================
# RECHERCHE ET TRAITEMENT DES FICHIERS 2025 ET 2026
# ==========================================
all_files = glob.glob("*.xlsx") + glob.glob("../*.xlsx") + glob.glob("*.xlsm") + glob.glob("../*.xlsm")
valid_files = [f for f in all_files if not os.path.basename(f).startswith('~$') and 'Nettoyee' not in f and 'NON RECU' not in os.path.basename(f).upper()]

# 1. Fichier 2025 : figé, priorise le fichier local Production 2025.xlsx
candidates_2025 = [f for f in valid_files if '2025' in os.path.basename(f)]
candidates_2025.sort(key=lambda f: (1 if os.path.dirname(f) == '' else 0, os.path.getmtime(f)), reverse=True)

# 2. Fichier 2026 : production active
candidates_2026 = [f for f in valid_files if '2025' not in os.path.basename(f)]
def sort_priority_2026(filepath):
    fname = os.path.basename(filepath).upper()
    priority = 0
    if 'ETAT DE PRODUCTION' in fname:
        priority = 3
    elif 'ARO' in fname:
        priority = 2
    elif 'PRODUCTION' in fname or 'CLASSEUR' in fname:
        priority = 1
    return (priority, os.path.getmtime(filepath))

candidates_2026.sort(key=sort_priority_2026, reverse=True)

if not candidates_2026:
    print("Erreur : Aucun fichier Excel 2026 trouvé !")
    exit(1)

file_2026 = candidates_2026[0]
print(f"Fichier 2026 sélectionné : {file_2026}")

records_2026 = parse_excel_file(file_2026, 2026, start_id=1)
print(f"-> 2026 : {len(records_2026)} affaires (dont {sum(1 for r in records_2026 if r['etatBordereau'] == 'REÇU')} signées, {sum(1 for r in records_2026 if r['isSansSuite'])} sans suite).")

records_2025 = []
if candidates_2025:
    file_2025 = candidates_2025[0]
    print(f"Fichier 2025 sélectionné : {file_2025}")
    records_2025 = parse_excel_file(file_2025, 2025, start_id=len(records_2026) + 1)
    print(f"-> 2025 : {len(records_2025)} affaires (dont {sum(1 for r in records_2025 if r['etatBordereau'] == 'REÇU')} signées, {sum(1 for r in records_2025 if r['isSansSuite'])} sans suite).")
else:
    print("Avertissement : Aucun fichier 2025 détecté.")

payload = {
    'records2026': records_2026,
    'records2025': records_2025,
    'records': records_2026 + records_2025,
    'cleaning_log': []
}

with open('cleaned_records.json', 'w', encoding='utf-8') as f:
    json.dump(payload, f, ensure_ascii=False, indent=2)

print(f"\nSauvegarde cleaned_records.json terminée avec succès !")

# Re-run generate_dashboard.py
if os.path.exists('generate_dashboard.py'):
    print("\nRégénération de index.html...")
    os.system("python generate_dashboard.py")
