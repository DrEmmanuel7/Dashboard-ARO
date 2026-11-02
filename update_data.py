import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import json
import re
import os
import glob

print("Recherche du fichier de production ARO le plus récent...")

# Search candidate files in current dir and parent Downloads dir
candidates = glob.glob("*.xlsx") + glob.glob("../*.xlsx") + glob.glob("*.xlsm") + glob.glob("../*.xlsm")
candidates = [f for f in candidates if not os.path.basename(f).startswith('~$') and not 'Nettoyee' in f and not 'NON RECU' in os.path.basename(f).upper()]

# Prioritize files with 'ETAT' or 'ARO' or 'PRODUCTION' in the name
def sort_priority(filepath):
    fname = os.path.basename(filepath).upper()
    priority = 0
    if 'ETAT DE PRODUCTION' in fname:
        priority = 3
    elif 'ARO' in fname:
        priority = 2
    elif 'PRODUCTION' in fname or 'CLASSEUR' in fname:
        priority = 1
    return (priority, os.path.getmtime(filepath))

candidates.sort(key=sort_priority, reverse=True)

if not candidates:
    print("Aucun fichier Excel trouvé !")
    exit(1)

aro_file = candidates[0]
print(f"Fichier détecté : {aro_file} (modifié le {datetime.fromtimestamp(os.path.getmtime(aro_file))})")

with zipfile.ZipFile(aro_file, 'r') as z:
    shared_strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        ss_xml = z.read('xl/sharedStrings.xml')
        ss_root = ET.fromstring(ss_xml)
        for si in ss_root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
            t = si.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
            shared_strings.append(t.text if t is not None else '')
            
    # Find sheet
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

# Find header row
header_row_num = None
header_col_map = {}
for r in rows:
    row_num = int(r.attrib.get('r'))
    vals = {}
    for c in r.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
        ref = c.attrib.get('r')
        col = ''.join([ch for ch in ref if ch.isalpha()])
        t_attr = c.attrib.get('t')
        v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        val = v.text if v is not None else ''
        if t_attr == 's' and val.isdigit():
            val = shared_strings[int(val)]
        vals[col] = str(val).strip()
        
    line_vals = [x.upper() for x in vals.values() if x]
    if any('DATE DE SOUSCRIPTION' in x for x in line_vals):
        header_row_num = row_num
        header_col_map = vals
        break

if not header_row_num:
    header_row_num = 1
    # Fallback to row 1
    for c in rows[0].iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
        ref = c.attrib.get('r')
        col = ''.join([ch for ch in ref if ch.isalpha()])
        t_attr = c.attrib.get('t')
        v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        val = v.text if v is not None else ''
        if t_attr == 's' and val.isdigit():
            val = shared_strings[int(val)]
        header_col_map[col] = str(val).strip()

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
    if 'ATLAS REINSURANCE' in up:
        return 'ATLAS REINSURANCE CONSULTANTS'
    if up in ['SBRE', 'SB RE']:
        return 'SB RE'
    if 'KLAPTON' in up:
        return 'KLAPTON AFRICA'
    if up == 'SUNU GROUP':
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

parsed_records = []
counter_id = 1

for r in rows:
    row_num = int(r.attrib.get('r'))
    if row_num <= header_row_num:
        continue
        
    row_dict = {}
    for c in r.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
        ref = c.attrib.get('r')
        col = ''.join([ch for ch in ref if ch.isalpha()])
        t_attr = c.attrib.get('t')
        v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        val = v.text if v is not None else ''
        if t_attr == 's' and val.isdigit():
            val = shared_strings[int(val)]
        h_name = header_col_map.get(col, col)
        row_dict[h_name] = val
        
    affaire = row_dict.get('AFFAIRES', '').strip()
    id_val = row_dict.get('N° id', '').strip()
    if not affaire and not id_val:
        continue
        
    date_souscr = excel_date_to_str(row_dict.get('DATE DE SOUSCRIPTION'))
    date_recept = excel_date_to_str(row_dict.get('DATE DE RECEPTION DU BORDEREAU'))
    
    delai = None
    if date_souscr and date_recept:
        try:
            d1 = datetime.strptime(date_souscr, '%Y-%m-%d')
            d2 = datetime.strptime(date_recept, '%Y-%m-%d')
            delai = (d2 - d1).days
        except:
            delai = None

    cap100 = to_float(row_dict.get('CAPITAUX 100 %'))
    prim100 = to_float(row_dict.get('PRIMES 100 %'))
    capSCG = to_float(row_dict.get('CAPITAUX SCG-Ré'))
    primSCG = to_float(row_dict.get('PRIMES SCG-Ré'))
    comm = to_float(row_dict.get('COMMISSION'))
    taxe = to_float(row_dict.get('TAXE'))
    primNet = to_float(row_dict.get('PRIMES NETTES'))
    if primNet == 0.0 and primSCG > 0:
        primNet = primSCG - comm

    parsed_records.append({
        'id': counter_id,
        'dateSouscription': date_souscr,
        'dateReception': date_recept,
        'delaiReceptionJours': delai,
        'typeCession': row_dict.get('Type de Cession', 'CESSION COMMERCIALE FAC'),
        'affaire': affaire,
        'typeAffaire': row_dict.get("TYPE D'AFFAIRE", ''),
        'operations': row_dict.get('OPERATIONS', ''),
        'cedante': clean_cedante(row_dict.get('CEDANTE', '')),
        'apporteur': clean_apporteur(row_dict.get('APPORTEUR', '')),
        'typeApporteur': row_dict.get('TYPE APPORTEUR', ''),
        'pays': clean_pays(row_dict.get('PAYS (Lieu du risque)', '')),
        'bureau': row_dict.get('BUREAU', ''),
        'zoneAssurance': row_dict.get('ZONE ASSURANCE', ''),
        'zoneGeo': row_dict.get('ZONE GEOGRAPHIQUE', ''),
        'branche': clean_branche(row_dict.get('BRANCHES', '')),
        'garanties': row_dict.get('GARANTIES', ''),
        'typeAssurance': row_dict.get("TYPES D'ASSURANCE", ''),
        'dateEffet': excel_date_to_str(row_dict.get("DATE D'EFFET")),
        'dateEcheance': excel_date_to_str(row_dict.get("DATE D'ECHEANCE")),
        'devise': row_dict.get('DEVISES', 'XOF'),
        'tauxChange': to_float(row_dict.get('TAUX DE CHANGES', 1.0)),
        'tauxCession': to_float(row_dict.get('TAUX  DE CESSION')),
        'tauxCommission': to_float(row_dict.get('TAUX DE COMMISSION')),
        'capitaux100': cap100,
        'primes100': prim100,
        'capitauxSCG': capSCG,
        'primesSCG': primSCG,
        'commission': comm,
        'taxe': taxe,
        'primesNettes': primNet,
        'etatBordereau': 'REÇU' if date_recept else 'EN ATTENTE',
        'commentaires': row_dict.get('COMMENTAIRES ', '')
    })
    counter_id += 1

with open('cleaned_records.json', 'w', encoding='utf-8') as f:
    json.dump({'records': parsed_records, 'cleaning_log': []}, f, ensure_ascii=False, indent=2)

print(f"Chargement terminé : {len(parsed_records)} affaires standardisées.")

# Re-run generate_dashboard.py
os.system("python generate_dashboard.py")
print("Tableau de bord index.html régénéré avec succès !")
