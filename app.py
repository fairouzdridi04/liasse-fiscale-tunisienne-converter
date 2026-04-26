<<<<<<< HEAD
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import tempfile, os, re, zipfile
from datetime import datetime
from io import BytesIO
import os  

import webbrowser
import threading

app = Flask(__name__)
CORS(app)

# ── Constants ──────────────────────────────────────────────────────────────────

NS = "http://www.impots.finances.gov.tn/liasse"

# Formulaires complets selon cahier des charges (sans F6006)
FORMS = {
    "F6001": dict(tag="F6001", prefixes=["F60010","F60011","F60012","F60013"], start=1, end=68, xsd="F6001.xsd", name="Bilan - Actif"),
    "F6002": dict(tag="F6002", prefixes=["F60020","F60021"], start=1, end=53, xsd="F6002.xsd", name="Bilan - Passif"),
    "F6003": dict(tag="F6003", prefixes=["F60030","F60031"], start=2, end=89, xsd="F6003.xsd", name="État de résultat"),
    "F6004": dict(tag="F6004", prefixes=["F60040","F60041"], start=1, end=117, xsd="F6004.xsd", name="Flux de trésorerie"),
    "F6005": dict(tag="F6005", prefixes=["F60050","F60051"], start=1, end=108, xsd="F6005.xsd", name="Résultat fiscal"),
    "F6007": dict(tag="F6007", prefixes=["F60070","F60071","F60072","F60073"], start=1, end=10, xsd="F6007.xsd", name="Autres informations"),
}

# Sheet name pattern → form key
SHEET_FORM_MAP = [
    (r"information|info|informations",   "F6007"),
    (r"fiscal|rsultat|resultat fiscal",  "F6005"),
    (r"actif",                           "F6001"),
    (r"passif",                          "F6002"),
    (r"r[eé]sultat|compte de resultat",  "F6003"),
    (r"flux|tresorerie|et flux",         "F6004"),
]

# Mapping des codes numériques vers F6005
FISCAL_CODE_MAPPING = {
    "1": "F60050000",
    "2": "F60050001",
    "36": "F60050055",
    "37": "F60050056",
}

# ── Helpers ────────────────────────────────────────────────────────────────────

def to13(value):
    """Convertit une valeur en millimes sur 13 chiffres"""
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return "0000000000000"
    try:
        n = round(abs(float(value)) * 1000)
        return str(n).zfill(13)
    except Exception:
        return "0000000000000"

def esc(s):
    return (str(s or "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

def clean_value(val):
    """Nettoie une valeur : supprime espaces, remplace virgule par point"""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        cleaned = val.replace(" ", "").replace(",", ".")
        parts = cleaned.split(".")
        if len(parts) > 2:
            cleaned = "".join(parts[:-1]) + "." + parts[-1]
        try:
            return float(cleaned)
        except:
            return 0.0
    return 0.0

def all_codes_for_form(form_key):
    """Retourne la liste complète des codes pour un formulaire"""
    f = FORMS[form_key]
    codes = []
    for pfx in f["prefixes"]:
        for i in range(f["start"], f["end"] + 1):
            codes.append(f"{pfx}{i:03d}")
    return codes

def sheet_to_form(sheet_name, df=None):
    """Détecte quel formulaire correspond à une feuille (par nom ou par contenu)"""
    name = sheet_name.lower().strip()
    
    for pattern, form_key in SHEET_FORM_MAP:
        if re.search(pattern, name):
            print(f"✅ Détection par nom: {sheet_name} → {form_key}")
            return form_key
    
    if df is not None and len(df) > 0 and df.shape[1] >= 1:
        for idx in range(min(30, len(df))):
            try:
                code_raw = df.iloc[idx, 0]
                if code_raw is not None and pd.notna(code_raw):
                    code = str(code_raw).strip().upper()
                    if code.startswith("F6007"):
                        print(f"✅ Détection par contenu: {sheet_name} → F6007")
                        return "F6007"
            except:
                continue
    
    print(f"❌ Aucune détection pour: {sheet_name}")
    return None

def read_excel(file_obj, f6004_mode=None):
    """
    Parse le fichier Excel et extrait toutes les données
    f6004_mode: 'auto' pour le modèle AUTORISÉ (E flux auto), 'ref' pour le modèle RÉFÉRENCE (et flux ref)
    """
    xl = pd.ExcelFile(file_obj)

    # ── Header (feuille 'entet') ───────────────────────────────────────────────
    header = {}
    entet_sheet = next((n for n in xl.sheet_names if "entet" in n.lower()), None)
    if entet_sheet:
        df_e = xl.parse(entet_sheet, header=None)
        field_map = {
            "matricule":    "mat",
            "raison":       "rs",
            "nomet":        "rs",
            "nom":          "rs",
            "activite":     "act",
            "activité":     "act",
            "adresse":      "adr",
            "exercice":     "ex",
            "date debut":   "dd",
            "date cloture": "df",
            "Acte De Depot": "ad",
            "Nature Depot":  "nd",
        }
        for _, row in df_e.iterrows():
            lbl_raw = row.iloc[1] if len(row) > 1 else None
            val_raw = row.iloc[2] if len(row) > 2 else None
            if pd.isna(lbl_raw):
                continue
            lbl = str(lbl_raw).lower().strip()
            val = str(val_raw).strip() if pd.notna(val_raw) else ""
            for key, field in field_map.items():
                if key in lbl and field not in header:
                    header[field] = val

    # Valeurs par défaut
    now_year = datetime.now().year
    header.setdefault("mat", "")
    header.setdefault("rs",  "")
    header.setdefault("act", "")
    header.setdefault("adr", "")
    header.setdefault("ex",  str(now_year))
    header.setdefault("dd",  f"01/01/{now_year}")
    header.setdefault("df",  f"31/12/{now_year}")
    header.setdefault("nd",  "P ")
    header.setdefault("ad",  "0 ")

    # Conversion des dates (YYYY-MM-DD → DD/MM/YYYY)
    for date_field in ("dd", "df"):
        v = header[date_field]
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", v)
        if m:
            header[date_field] = f"{m.group(3)}/{m.group(2)}/{m.group(1)}"

    # ── Data sheets ────────────────────────────────────────────────────────────
    code_maps   = {k: {} for k in FORMS}
    label_maps  = {k: {} for k in FORMS}
    sheets_used = {k: [] for k in FORMS}

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)
        form_key = sheet_to_form(sheet_name, df)
        
        if form_key is None:
            continue
        
        # Pour F6004, filtrer selon le mode sélectionné
        if form_key == "F6004" and f6004_mode:
            sheet_lower = sheet_name.lower()
            if f6004_mode == "auto" and "e flux auto" not in sheet_lower:
                print(f"⏭️ Ignoré (mode AUTO): {sheet_name}")
                continue
            elif f6004_mode == "ref" and "et flux ref" not in sheet_lower:
                print(f"⏭️ Ignoré (mode REF): {sheet_name}")
                continue

        sheets_used[form_key].append(sheet_name)

        for _, row in df.iterrows():
            code_raw = row.iloc[0] if len(row) > 0 else None
            label = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
            val = row.iloc[2] if len(row) > 2 else None

            if pd.isna(code_raw):
                continue

            code = str(code_raw).strip()
            float_val = clean_value(val)

            # Pour F6005 : codes numériques
            if form_key == "F6005":
                clean_code = code.replace(".0", "") if code.endswith(".0") else code
                if clean_code.isdigit() and clean_code in FISCAL_CODE_MAPPING:
                    f6005_code = FISCAL_CODE_MAPPING[clean_code]
                    if f6005_code in code_maps[form_key]:
                        if abs(float_val) > abs(code_maps[form_key][f6005_code]):
                            code_maps[form_key][f6005_code] = float_val
                            label_maps[form_key][f6005_code] = label
                    else:
                        code_maps[form_key][f6005_code] = float_val
                        label_maps[form_key][f6005_code] = label
                elif re.match(r"^F6005\d{5}$", code):
                    if code in code_maps[form_key]:
                        if abs(float_val) > abs(code_maps[form_key][code]):
                            code_maps[form_key][code] = float_val
                            label_maps[form_key][code] = label
                    else:
                        code_maps[form_key][code] = float_val
                        label_maps[form_key][code] = label

            # Pour F6007 : codes F6007xxxx
            elif form_key == "F6007" and re.match(r"^F6007\d{5}$", code):
                if code in code_maps[form_key]:
                    if abs(float_val) > abs(code_maps[form_key][code]):
                        code_maps[form_key][code] = float_val
                        label_maps[form_key][code] = label
                else:
                    code_maps[form_key][code] = float_val
                    label_maps[form_key][code] = label

            # Pour les autres formulaires F6001-F6004
            elif re.match(r"^F6\d{7}$", code):
                expected_prefixes = FORMS[form_key]["prefixes"]
                if any(code.startswith(p) for p in expected_prefixes):
                    if code in code_maps[form_key]:
                        if abs(float_val) > abs(code_maps[form_key][code]):
                            code_maps[form_key][code] = float_val
                            label_maps[form_key][code] = label
                    else:
                        code_maps[form_key][code] = float_val
                        label_maps[form_key][code] = label

    return header, code_maps, label_maps, sheets_used

def build_xml(form_key, header, code_map):
    """Construit le XML pour un formulaire donné"""
    f      = FORMS[form_key]
    tag    = f["tag"]
    SCHEMA = f"{NS}/{f['xsd']}"

    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        f'<lf:{tag} xmlns:lf="{NS}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="{NS} {SCHEMA}">',
        '  <lf:VersionDocument>1.0</lf:VersionDocument>',
        '  <lf:Entete>',
        f'    <lf:MatriculeFiscalDeclarant>{esc(header.get("mat",""))}</lf:MatriculeFiscalDeclarant>',
        f'    <lf:NometPrenomouRaisonSociale>{esc(header.get("rs",""))}</lf:NometPrenomouRaisonSociale>',
        f'    <lf:Activite>{esc(header.get("act",""))}</lf:Activite>',
        f'    <lf:Adresse>{esc(header.get("adr",""))}</lf:Adresse>',
        f'    <lf:Exercice>{esc(header.get("ex",""))}</lf:Exercice>',
        f'    <lf:DateDebutExercice>{esc(header.get("dd",""))}</lf:DateDebutExercice>',
        f'    <lf:DateClotureExercice>{esc(header.get("df",""))}</lf:DateClotureExercice>',
        f'    <lf:ActeDeDepot>{esc(header.get("ad","0 "))}</lf:ActeDeDepot>',
        f'    <lf:NatureDepot>{esc(header.get("nd","P "))}</lf:NatureDepot>',
        '  </lf:Entete>',
        '  <lf:Details>',
    ]

    for code in all_codes_for_form(form_key):
        val = code_map.get(code, 0)
        lines.append(f'    <lf:{code}>{to13(val)}</lf:{code}>')

    lines += [f'  </lf:Details>', f'</lf:{tag}>']
    return "\n".join(lines)

# ── Routes Flask ───────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('dashboard.html')

@app.route('/preview', methods=['POST'])
def preview():
    """Retourne un résumé des formulaires détectés"""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier'}), 400
    try:
        # Pour l'aperçu, on ne filtre pas encore les modèles F6004
        header, code_maps, label_maps, sheets_used = read_excel(request.files['file'])
        forms_summary = {}
        for fk in FORMS:
            codes = all_codes_for_form(fk)
            non_zero = sum(1 for c in codes if code_maps[fk].get(c, 0) != 0)
            forms_summary[fk] = {
                "total": len(codes),
                "non_zero": non_zero,
                "sheets": sheets_used[fk],
                "detected": bool(sheets_used[fk]),
                "name": FORMS[fk]["name"]
            }
        return jsonify({"header": header, "forms": forms_summary})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/convert', methods=['POST'])
def convert():
    """Génère les fichiers XML (un seul ou ZIP)"""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier'}), 400

    try:
        file = request.files['file']
        f6004_mode = request.form.get('f6004_mode', None)  # Récupérer le mode sélectionné
        
        print(f"📌 Mode F6004 sélectionné: {f6004_mode if f6004_mode else 'auto-détection'}")
        
        header, code_maps, label_maps, sheets_used = read_excel(file, f6004_mode)

        # Surcharge par les champs du formulaire
        for field in ("mat", "rs", "act", "adr", "ex", "dd", "df", "nd", "ad"):
            v = request.form.get(field, "").strip()
            if v:
                header[field] = v

        # Formulaires à exporter
        requested = request.form.get("form_keys", "all")
        if requested == "all":
            form_keys = [fk for fk in FORMS if sheets_used[fk]]
        else:
            form_keys = [fk.strip() for fk in requested.split(",") if fk.strip() in FORMS]

        if not form_keys:
            return jsonify({'error': 'Aucune feuille reconnue dans ce fichier.'}), 400

        mat_clean = (header.get("mat", "") or "INCONNU").replace(" ", "")
        ex_clean = header.get("ex", str(datetime.now().year))

        # Un seul formulaire → XML direct
        if len(form_keys) == 1:
            fk = form_keys[0]
            xml = build_xml(fk, header, code_maps[fk])
            tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False, encoding="utf-8")
            tmp.write(xml)
            tmp.close()
            fname = f"{fk}-{mat_clean}-{ex_clean}.xml"
            return send_file(tmp.name, as_attachment=True, download_name=fname, mimetype="application/xml")

        # Plusieurs formulaires → ZIP
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fk in form_keys:
                xml = build_xml(fk, header, code_maps[fk])
                fname = f"{fk}-{mat_clean}-{ex_clean}.xml"
                zf.writestr(fname, xml.encode("utf-8"))
        zip_buf.seek(0)

        zip_name = f"LiasseFiscale-{mat_clean}-{ex_clean}.zip"
        return send_file(zip_buf, as_attachment=True, download_name=zip_name, mimetype="application/zip")

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/template')
def download_template():
    """Télécharge le template Excel statique"""
    import sys
    import os
    
    # Détecter si on est dans un exe ou en développement
    if getattr(sys, 'frozen', False):
        # On est dans un exe
        base_path = sys._MEIPASS
    else:
        # On est en développement
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    # Chercher le template
    template_path = os.path.join(base_path, 'static', 'template_liasse_fiscale.xlsx')
    
    # Vérifier si le fichier existe
    if not os.path.exists(template_path):
        return jsonify({'error': 'Template non trouvé. Vérifiez que le fichier static/template_liasse_fiscale.xlsx existe'}), 404
    
    return send_file(
        template_path,
        as_attachment=True,
        download_name='template_liasse_fiscale.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == "__main__":
    print("\n" + "="*70)
    print("  📊  LIASSE FISCALE TUNISIENNE — Convertisseur F6001 à F6007")
    print("="*70)
    print("  🚀  Serveur démarré sur http://127.0.0.1:5000")
    print("  📋  Formulaires supportés: F6001 · F6002 · F6003 · F6004 · F6005 · F6007")
    print("  🔄  F6004 - Deux modèles disponibles")
    print("  ⚠   Pour arrêter : CTRL+C")
    print("="*70 + "\n")
    
    # Ouvrir le navigateur automatiquement après 1 seconde
    def open_browser():
        webbrowser.open('http://127.0.0.1:5000')
    
    threading.Timer(1, open_browser).start()
    
=======
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import tempfile, os, re, zipfile
from datetime import datetime
from io import BytesIO
import os  

import webbrowser
import threading

app = Flask(__name__)
CORS(app)

# ── Constants ──────────────────────────────────────────────────────────────────

NS = "http://www.impots.finances.gov.tn/liasse"

# Formulaires complets selon cahier des charges (sans F6006)
FORMS = {
    "F6001": dict(tag="F6001", prefixes=["F60010","F60011","F60012","F60013"], start=1, end=68, xsd="F6001.xsd", name="Bilan - Actif"),
    "F6002": dict(tag="F6002", prefixes=["F60020","F60021"], start=1, end=53, xsd="F6002.xsd", name="Bilan - Passif"),
    "F6003": dict(tag="F6003", prefixes=["F60030","F60031"], start=2, end=89, xsd="F6003.xsd", name="État de résultat"),
    "F6004": dict(tag="F6004", prefixes=["F60040","F60041"], start=1, end=117, xsd="F6004.xsd", name="Flux de trésorerie"),
    "F6005": dict(tag="F6005", prefixes=["F60050","F60051"], start=1, end=108, xsd="F6005.xsd", name="Résultat fiscal"),
    "F6007": dict(tag="F6007", prefixes=["F60070","F60071","F60072","F60073"], start=1, end=10, xsd="F6007.xsd", name="Autres informations"),
}

# Sheet name pattern → form key
SHEET_FORM_MAP = [
    (r"information|info|informations",   "F6007"),
    (r"fiscal|rsultat|resultat fiscal",  "F6005"),
    (r"actif",                           "F6001"),
    (r"passif",                          "F6002"),
    (r"r[eé]sultat|compte de resultat",  "F6003"),
    (r"flux|tresorerie|et flux",         "F6004"),
]

# Mapping des codes numériques vers F6005
FISCAL_CODE_MAPPING = {
    "1": "F60050000",
    "2": "F60050001",
    "36": "F60050055",
    "37": "F60050056",
}

# ── Helpers ────────────────────────────────────────────────────────────────────

def to13(value):
    """Convertit une valeur en millimes sur 13 chiffres"""
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return "0000000000000"
    try:
        n = round(abs(float(value)) * 1000)
        return str(n).zfill(13)
    except Exception:
        return "0000000000000"

def esc(s):
    return (str(s or "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

def clean_value(val):
    """Nettoie une valeur : supprime espaces, remplace virgule par point"""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        cleaned = val.replace(" ", "").replace(",", ".")
        parts = cleaned.split(".")
        if len(parts) > 2:
            cleaned = "".join(parts[:-1]) + "." + parts[-1]
        try:
            return float(cleaned)
        except:
            return 0.0
    return 0.0

def all_codes_for_form(form_key):
    """Retourne la liste complète des codes pour un formulaire"""
    f = FORMS[form_key]
    codes = []
    for pfx in f["prefixes"]:
        for i in range(f["start"], f["end"] + 1):
            codes.append(f"{pfx}{i:03d}")
    return codes

def sheet_to_form(sheet_name, df=None):
    """Détecte quel formulaire correspond à une feuille (par nom ou par contenu)"""
    name = sheet_name.lower().strip()
    
    for pattern, form_key in SHEET_FORM_MAP:
        if re.search(pattern, name):
            print(f"✅ Détection par nom: {sheet_name} → {form_key}")
            return form_key
    
    if df is not None and len(df) > 0 and df.shape[1] >= 1:
        for idx in range(min(30, len(df))):
            try:
                code_raw = df.iloc[idx, 0]
                if code_raw is not None and pd.notna(code_raw):
                    code = str(code_raw).strip().upper()
                    if code.startswith("F6007"):
                        print(f"✅ Détection par contenu: {sheet_name} → F6007")
                        return "F6007"
            except:
                continue
    
    print(f"❌ Aucune détection pour: {sheet_name}")
    return None

def read_excel(file_obj, f6004_mode=None):
    """
    Parse le fichier Excel et extrait toutes les données
    f6004_mode: 'auto' pour le modèle AUTORISÉ (E flux auto), 'ref' pour le modèle RÉFÉRENCE (et flux ref)
    """
    xl = pd.ExcelFile(file_obj)

    # ── Header (feuille 'entet') ───────────────────────────────────────────────
    header = {}
    entet_sheet = next((n for n in xl.sheet_names if "entet" in n.lower()), None)
    if entet_sheet:
        df_e = xl.parse(entet_sheet, header=None)
        field_map = {
            "matricule":    "mat",
            "raison":       "rs",
            "nomet":        "rs",
            "nom":          "rs",
            "activite":     "act",
            "activité":     "act",
            "adresse":      "adr",
            "exercice":     "ex",
            "date debut":   "dd",
            "date cloture": "df",
            "Acte De Depot": "ad",
            "Nature Depot":  "nd",
        }
        for _, row in df_e.iterrows():
            lbl_raw = row.iloc[1] if len(row) > 1 else None
            val_raw = row.iloc[2] if len(row) > 2 else None
            if pd.isna(lbl_raw):
                continue
            lbl = str(lbl_raw).lower().strip()
            val = str(val_raw).strip() if pd.notna(val_raw) else ""
            for key, field in field_map.items():
                if key in lbl and field not in header:
                    header[field] = val

    # Valeurs par défaut
    now_year = datetime.now().year
    header.setdefault("mat", "")
    header.setdefault("rs",  "")
    header.setdefault("act", "")
    header.setdefault("adr", "")
    header.setdefault("ex",  str(now_year))
    header.setdefault("dd",  f"01/01/{now_year}")
    header.setdefault("df",  f"31/12/{now_year}")
    header.setdefault("nd",  "P ")
    header.setdefault("ad",  "0 ")

    # Conversion des dates (YYYY-MM-DD → DD/MM/YYYY)
    for date_field in ("dd", "df"):
        v = header[date_field]
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", v)
        if m:
            header[date_field] = f"{m.group(3)}/{m.group(2)}/{m.group(1)}"

    # ── Data sheets ────────────────────────────────────────────────────────────
    code_maps   = {k: {} for k in FORMS}
    label_maps  = {k: {} for k in FORMS}
    sheets_used = {k: [] for k in FORMS}

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)
        form_key = sheet_to_form(sheet_name, df)
        
        if form_key is None:
            continue
        
        # Pour F6004, filtrer selon le mode sélectionné
        if form_key == "F6004" and f6004_mode:
            sheet_lower = sheet_name.lower()
            if f6004_mode == "auto" and "e flux auto" not in sheet_lower:
                print(f"⏭️ Ignoré (mode AUTO): {sheet_name}")
                continue
            elif f6004_mode == "ref" and "et flux ref" not in sheet_lower:
                print(f"⏭️ Ignoré (mode REF): {sheet_name}")
                continue

        sheets_used[form_key].append(sheet_name)

        for _, row in df.iterrows():
            code_raw = row.iloc[0] if len(row) > 0 else None
            label = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
            val = row.iloc[2] if len(row) > 2 else None

            if pd.isna(code_raw):
                continue

            code = str(code_raw).strip()
            float_val = clean_value(val)

            # Pour F6005 : codes numériques
            if form_key == "F6005":
                clean_code = code.replace(".0", "") if code.endswith(".0") else code
                if clean_code.isdigit() and clean_code in FISCAL_CODE_MAPPING:
                    f6005_code = FISCAL_CODE_MAPPING[clean_code]
                    if f6005_code in code_maps[form_key]:
                        if abs(float_val) > abs(code_maps[form_key][f6005_code]):
                            code_maps[form_key][f6005_code] = float_val
                            label_maps[form_key][f6005_code] = label
                    else:
                        code_maps[form_key][f6005_code] = float_val
                        label_maps[form_key][f6005_code] = label
                elif re.match(r"^F6005\d{5}$", code):
                    if code in code_maps[form_key]:
                        if abs(float_val) > abs(code_maps[form_key][code]):
                            code_maps[form_key][code] = float_val
                            label_maps[form_key][code] = label
                    else:
                        code_maps[form_key][code] = float_val
                        label_maps[form_key][code] = label

            # Pour F6007 : codes F6007xxxx
            elif form_key == "F6007" and re.match(r"^F6007\d{5}$", code):
                if code in code_maps[form_key]:
                    if abs(float_val) > abs(code_maps[form_key][code]):
                        code_maps[form_key][code] = float_val
                        label_maps[form_key][code] = label
                else:
                    code_maps[form_key][code] = float_val
                    label_maps[form_key][code] = label

            # Pour les autres formulaires F6001-F6004
            elif re.match(r"^F6\d{7}$", code):
                expected_prefixes = FORMS[form_key]["prefixes"]
                if any(code.startswith(p) for p in expected_prefixes):
                    if code in code_maps[form_key]:
                        if abs(float_val) > abs(code_maps[form_key][code]):
                            code_maps[form_key][code] = float_val
                            label_maps[form_key][code] = label
                    else:
                        code_maps[form_key][code] = float_val
                        label_maps[form_key][code] = label

    return header, code_maps, label_maps, sheets_used

def build_xml(form_key, header, code_map):
    """Construit le XML pour un formulaire donné"""
    f      = FORMS[form_key]
    tag    = f["tag"]
    SCHEMA = f"{NS}/{f['xsd']}"

    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        f'<lf:{tag} xmlns:lf="{NS}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="{NS} {SCHEMA}">',
        '  <lf:VersionDocument>1.0</lf:VersionDocument>',
        '  <lf:Entete>',
        f'    <lf:MatriculeFiscalDeclarant>{esc(header.get("mat",""))}</lf:MatriculeFiscalDeclarant>',
        f'    <lf:NometPrenomouRaisonSociale>{esc(header.get("rs",""))}</lf:NometPrenomouRaisonSociale>',
        f'    <lf:Activite>{esc(header.get("act",""))}</lf:Activite>',
        f'    <lf:Adresse>{esc(header.get("adr",""))}</lf:Adresse>',
        f'    <lf:Exercice>{esc(header.get("ex",""))}</lf:Exercice>',
        f'    <lf:DateDebutExercice>{esc(header.get("dd",""))}</lf:DateDebutExercice>',
        f'    <lf:DateClotureExercice>{esc(header.get("df",""))}</lf:DateClotureExercice>',
        f'    <lf:ActeDeDepot>{esc(header.get("ad","0 "))}</lf:ActeDeDepot>',
        f'    <lf:NatureDepot>{esc(header.get("nd","P "))}</lf:NatureDepot>',
        '  </lf:Entete>',
        '  <lf:Details>',
    ]

    for code in all_codes_for_form(form_key):
        val = code_map.get(code, 0)
        lines.append(f'    <lf:{code}>{to13(val)}</lf:{code}>')

    lines += [f'  </lf:Details>', f'</lf:{tag}>']
    return "\n".join(lines)

# ── Routes Flask ───────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('dashboard.html')

@app.route('/preview', methods=['POST'])
def preview():
    """Retourne un résumé des formulaires détectés"""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier'}), 400
    try:
        # Pour l'aperçu, on ne filtre pas encore les modèles F6004
        header, code_maps, label_maps, sheets_used = read_excel(request.files['file'])
        forms_summary = {}
        for fk in FORMS:
            codes = all_codes_for_form(fk)
            non_zero = sum(1 for c in codes if code_maps[fk].get(c, 0) != 0)
            forms_summary[fk] = {
                "total": len(codes),
                "non_zero": non_zero,
                "sheets": sheets_used[fk],
                "detected": bool(sheets_used[fk]),
                "name": FORMS[fk]["name"]
            }
        return jsonify({"header": header, "forms": forms_summary})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/convert', methods=['POST'])
def convert():
    """Génère les fichiers XML (un seul ou ZIP)"""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier'}), 400

    try:
        file = request.files['file']
        f6004_mode = request.form.get('f6004_mode', None)  # Récupérer le mode sélectionné
        
        print(f"📌 Mode F6004 sélectionné: {f6004_mode if f6004_mode else 'auto-détection'}")
        
        header, code_maps, label_maps, sheets_used = read_excel(file, f6004_mode)

        # Surcharge par les champs du formulaire
        for field in ("mat", "rs", "act", "adr", "ex", "dd", "df", "nd", "ad"):
            v = request.form.get(field, "").strip()
            if v:
                header[field] = v

        # Formulaires à exporter
        requested = request.form.get("form_keys", "all")
        if requested == "all":
            form_keys = [fk for fk in FORMS if sheets_used[fk]]
        else:
            form_keys = [fk.strip() for fk in requested.split(",") if fk.strip() in FORMS]

        if not form_keys:
            return jsonify({'error': 'Aucune feuille reconnue dans ce fichier.'}), 400

        mat_clean = (header.get("mat", "") or "INCONNU").replace(" ", "")
        ex_clean = header.get("ex", str(datetime.now().year))

        # Un seul formulaire → XML direct
        if len(form_keys) == 1:
            fk = form_keys[0]
            xml = build_xml(fk, header, code_maps[fk])
            tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False, encoding="utf-8")
            tmp.write(xml)
            tmp.close()
            fname = f"{fk}-{mat_clean}-{ex_clean}.xml"
            return send_file(tmp.name, as_attachment=True, download_name=fname, mimetype="application/xml")

        # Plusieurs formulaires → ZIP
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fk in form_keys:
                xml = build_xml(fk, header, code_maps[fk])
                fname = f"{fk}-{mat_clean}-{ex_clean}.xml"
                zf.writestr(fname, xml.encode("utf-8"))
        zip_buf.seek(0)

        zip_name = f"LiasseFiscale-{mat_clean}-{ex_clean}.zip"
        return send_file(zip_buf, as_attachment=True, download_name=zip_name, mimetype="application/zip")

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/template')
def download_template():
    """Télécharge le template Excel statique"""
    import sys
    import os
    
    # Détecter si on est dans un exe ou en développement
    if getattr(sys, 'frozen', False):
        # On est dans un exe
        base_path = sys._MEIPASS
    else:
        # On est en développement
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    # Chercher le template
    template_path = os.path.join(base_path, 'static', 'template_liasse_fiscale.xlsx')
    
    # Vérifier si le fichier existe
    if not os.path.exists(template_path):
        return jsonify({'error': 'Template non trouvé. Vérifiez que le fichier static/template_liasse_fiscale.xlsx existe'}), 404
    
    return send_file(
        template_path,
        as_attachment=True,
        download_name='template_liasse_fiscale.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == "__main__":
    print("\n" + "="*70)
    print("  📊  LIASSE FISCALE TUNISIENNE — Convertisseur F6001 à F6007")
    print("="*70)
    print("  🚀  Serveur démarré sur http://127.0.0.1:5000")
    print("  📋  Formulaires supportés: F6001 · F6002 · F6003 · F6004 · F6005 · F6007")
    print("  🔄  F6004 - Deux modèles disponibles")
    print("  ⚠   Pour arrêter : CTRL+C")
    print("="*70 + "\n")
    
    # Ouvrir le navigateur automatiquement après 1 seconde
    def open_browser():
        webbrowser.open('http://127.0.0.1:5000')
    
    threading.Timer(1, open_browser).start()
    
>>>>>>> master
    app.run(debug=False, host="127.0.0.1", port=5000)