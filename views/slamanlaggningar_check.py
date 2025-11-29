from pathlib import Path
from typing import List
import re
import pandas as pd

from flask import Blueprint, request, flash, redirect, url_for, session
from utils.file_utils import allowed_file, create_session_paths, cleanup_folder, UPLOAD_FOLDER

bp = Blueprint('slamanlaggningar_check', __name__)


def _norm(s: str) -> str:
    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def expected_count_from_freq(freq: str):

    if pd.isna(freq):
        return None
    s = str(freq).strip().lower()
    if 'vartannat år' in s or 'vartannat-år' in s:
        return 1

    m = re.search(r'(\d+)', s)
    if m:
        try:
            val = int(m.group(1))
            if 1 <= val <= 12:
                return val
        except Exception:
            return None
    return None


def extract_week_tokens(s: str) -> List[str]:

    if pd.isna(s):
        return []
    s_norm = str(s).lower()
    matches = re.findall(r'vecka\s*\d{1,2}', s_norm)
    return [m.strip() for m in matches]


def process_slamanlaggningar(input_path: Path, output_path: Path) -> int:

    df = pd.read_excel(input_path)

    required_cols = {
        'Affärsenhet',
        'Kundnr',
        'Flexplatsadress',
        'Flextjänstnr',
        'Flexgrupp namn',
        'Flextyp',
        'Utförandeområde flextjänst',
        'Utförandeområde flexplats',
        'Hämtfrekvens',
        'Ind. körtursplan',
        'Körtursnamn'
    }
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Saknar kolumner: {', '.join(missing)}")

    results: List[dict] = []
    row_level_flags = []  # samla index för rader som redan är avvikande (för att undvika dubbletter)

    for idx, row in df.iterrows():
        orsaker = []

        ua_tjanst = row.get('Utförandeområde flextjänst')
        ua_plats = row.get('Utförandeområde flexplats')
        ind_kort = row.get('Ind. körtursplan')
        kortnamn = row.get('Körtursnamn')
        hamt = row.get('Hämtfrekvens')
        flextyp = row.get('Flextyp')

        # Om "Utförandeområde flexplats" eller "Utförandeområde flextjänst" saknas
        if pd.isna(ua_tjanst) or str(ua_tjanst).strip() == "":
            orsaker.append("Saknar Utförandeområde flextjänst")
        if pd.isna(ua_plats) or str(ua_plats).strip() == "":
            orsaker.append("Saknar Utförandeområde flexplats")
        # Om båda finns men mismatch
        if (not pd.isna(ua_tjanst) and not pd.isna(ua_plats)) and (_norm(ua_tjanst) != _norm(ua_plats)):
            orsaker.append("Utförandeområde flextjänst ≠ Utförandeområde flexplats")

        # Ind. körtursplan eller Körtursnamn saknas
        if pd.isna(ind_kort) or str(ind_kort).strip() == "":
            orsaker.append("Saknar Ind. körtursplan")
        if pd.isna(kortnamn) or str(kortnamn).strip() == "":
            orsaker.append("Saknar Körtursnamn")

        # Vartannat år, då måste Ind. körtursplan måste innehålla 'udda år' eller 'jämna år'
        if isinstance(hamt, str) or not pd.isna(hamt):
            hamt_norm = _norm(hamt)
            if 'vartannat år' in hamt_norm:
                if ind_kort is None or not (('udda år' in _norm(ind_kort)) or ('jämna år' in _norm(ind_kort))):
                    orsaker.append("Hämtfrekvens 'Vartannat år' kräver 'udda år' eller 'jämna år' i Ind. körtursplan")

        # Veckonummer i Ind. körtursplan måste finnas i Körtursnamn
        if not (pd.isna(ind_kort) or str(ind_kort).strip() == ""):
            week_tokens = extract_week_tokens(ind_kort)
            for wk in week_tokens:
                if wk not in _norm(kortnamn):
                    orsaker.append(f"Ind. körtursplan innehåller '{wk}' men Körtursnamn innehåller inte '{wk}'")

        # Bud-regeln (gäller åt båda håll)
        hamt_norm = _norm(hamt)
        ind_norm = _norm(ind_kort)
        kort_norm = _norm(kortnamn)

        # Om Hämtfrekvens innehåller 'bud'
        if 'bud' in hamt_norm:
            if 'bud' not in ind_norm:
                orsaker.append("Hämtfrekvens 'Bud' kräver 'Budning' i Ind. körtursplan")
            if 'bud' not in kort_norm:
                orsaker.append("Hämtfrekvens 'Bud' kräver 'bud' i Körtursnamn")
        # Om Ind. körtursplan innehåller budning
        if 'bud' in ind_norm:
            if 'bud' not in hamt_norm:
                orsaker.append("Ind. körtursplan 'Budning' kräver Hämtfrekvens 'Bud'")
            if 'bud' not in kort_norm:
                orsaker.append("Ind. körtursplan 'Budning' kräver 'bud' i Körtursnamn")
        # Om Körtursnamn innehåller bud
        if 'bud' in kort_norm:
            if 'bud' not in hamt_norm and 'bud' not in ind_norm:
                orsaker.append("Körtursnamn innehåller 'bud' men saknar Bud i Hämtfrekvens/Ind. körtursplan")

        # Hämtfrekvens får ej vara tom
        if pd.isna(hamt) or str(hamt).strip() == "":
            orsaker.append("Saknar Hämtfrekvens")

        if orsaker:
            row_level_flags.append(idx)
            results.append({
                "Affärsenhet": row.get('Affärsenhet'),
                "Kundnr": row.get('Kundnr'),
                "Flexplatsadress": row.get('Flexplatsadress'),
                "Flextjänstnr": row.get('Flextjänstnr'),
                "Flexgrupp namn": row.get('Flexgrupp namn'),
                "Flextyp": row.get('Flextyp'),
                "Utförandeområde flextjänst": ua_tjanst,
                "Utförandeområde flexplats": ua_plats,
                "Hämtfrekvens": hamt,
                "Ind. körtursplan": ind_kort,
                "Körtursnamn": kortnamn,
                "Orsak": "; ".join(orsaker)
            })

    # Räkna förekomster per Flextjänstnr
    counts = df['Flextjänstnr'].value_counts().to_dict()

    # För varje unikt flextjänstnr, bestäm förväntat antal från raderna i gruppen
    for flextnr, cnt in counts.items():
        grp = df[df['Flextjänstnr'] == flextnr]
        # Hämta frekvensvärden i gruppen (unika)
        freqs = [f for f in grp['Hämtfrekvens'].dropna().unique()]
        expected_vals = sorted({expected_count_from_freq(f) for f in freqs if expected_count_from_freq(f) is not None})
        # Om flera olika förväntade värden finns i samma grupp genererar det avvikelse
        if len(expected_vals) == 0:
            continue
        if len(expected_vals) > 1:
            rep = grp.iloc[0]
            results.append({
                "Affärsenhet": rep.get('Affärsenhet'),
                "Kundnr": rep.get('Kundnr'),
                "Flexplatsadress": rep.get('Flexplatsadress'),
                "Flextjänstnr": flextnr,
                "Flexgrupp namn": rep.get('Flexgrupp namn'),
                "Flextyp": rep.get('Flextyp'),
                "Utförandeområde flextjänst": rep.get('Utförandeområde flextjänst'),
                "Utförandeområde flexplats": rep.get('Utförandeområde flexplats'),
                "Hämtfrekvens": ", ".join(map(str, freqs)),
                "Ind. körtursplan": rep.get('Ind. körtursplan'),
                "Körtursnamn": rep.get('Körtursnamn'),
                "Orsak": f"Inkonsekventa Hämtfrekvenser inom flextjänst (ger förväntningar {expected_vals})"
            })
            continue

        expected = expected_vals[0]
        if expected is None:
            continue

        if cnt != expected:
            # Fel på antal förekomster av flextjänstnr i relation till förväntat från hämtfrekvens genererar avvikelse
            rep = grp.iloc[0]
            results.append({
                "Affärsenhet": rep.get('Affärsenhet'),
                "Kundnr": rep.get('Kundnr'),
                "Flexplatsadress": rep.get('Flexplatsadress'),
                "Flextjänstnr": flextnr,
                "Flexgrupp namn": rep.get('Flexgrupp namn'),
                "Flextyp": rep.get('Flextyp'),
                "Utförandeområde flextjänst": rep.get('Utförandeområde flextjänst'),
                "Utförandeområde flexplats": rep.get('Utförandeområde flexplats'),
                "Hämtfrekvens": rep.get('Hämtfrekvens'),
                "Ind. körtursplan": rep.get('Ind. körtursplan'),
                "Körtursnamn": rep.get('Körtursnamn'),
                "Orsak": f"Flextjänstnr förekommer {cnt} gånger men förväntat {expected} enligt Hämtfrekvens"
            })

    out_df = pd.DataFrame(results)

    # Skriv resultat till Excel
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Avvikelser")
        workbook = writer.book
        worksheet = writer.sheets["Avvikelser"]

        header_fmt = workbook.add_format({"align": "left", "bold": True})
        for col_idx, value in enumerate(out_df.columns):
            worksheet.write(0, col_idx, value, header_fmt)

        cell_fmt = workbook.add_format({"align": "left"})
        worksheet.set_column(0, len(out_df.columns) - 1, 30, cell_fmt)

    return len(out_df)


# Endpoint för filuppladdning och bearbetning
@bp.route('/upload/slamanlaggningar_check', methods=['POST'])
def slamanlaggningar_upload():
    if 'file' not in request.files:
        flash('Ingen fil i anropet')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('Du måste välja en fil')
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash('Endast Excel-filer (.xlsx) tillåtna')
        return redirect(url_for('index'))

    cleanup_folder()

    input_path, session_id = create_session_paths(file.filename)
    output_filename = f"avvikelser_slamanlaggningar_{session_id}.xlsx"
    output_path = UPLOAD_FOLDER / output_filename

    file.save(input_path)

    try:
        deviations = process_slamanlaggningar(input_path, output_path)
    except ValueError as e:
        flash(str(e))
        return redirect(url_for('index'))
    except Exception:
        flash('Fel vid bearbetning av filen')
        return redirect(url_for('index'))

    if deviations > 0:
        message = f"{deviations} anläggningar har avvikelser som behöver hanteras"
    else:
        message = "Inga avvikelser hittades."

    session_key = f"result_{session_id}"
    session[session_key] = {
        "deviations": deviations,
        "output_filename": output_filename,
        "back": "index",
        "message": message
    }

    return redirect(url_for('success', file=session_id))
