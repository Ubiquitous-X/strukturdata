from pathlib import Path
from typing import List
import pandas as pd

from flask import Blueprint, request, flash, redirect, url_for, session
from utils.file_utils import allowed_file, create_session_paths, cleanup_folder, UPLOAD_FOLDER

bp = Blueprint('antalsvarde_individer', __name__)


def process_karl(input_path: Path, output_path: Path) -> int:

    df = pd.read_excel(input_path)

    # Kontrollera obligatoriska kolumner
    required_cols = {
        'Affärsenhet',
        'Status',
        'Flexplatsadress',
        'Flextjänstnr',
        'Fraktion',
        'Flextyp',
        'Extern referens',
        'Antal kärl'
    }
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Saknar kolumner: {', '.join(missing)}")

    # Hjälpfunktion: konvertera antal kärl till int eller None
    def to_int_or_none(x):
        if pd.isna(x):
            return None
        s = str(x).strip()
        if s == "":
            return None
        try:
            return int(float(s))
        except Exception:
            return None

    df['_antal_karl_num'] = df['Antal kärl'].apply(to_int_or_none)
    df['_extern_ref'] = df['Extern referens'].astype(str).str.strip()
    df['_extern_missing'] = df['_extern_ref'].isin(['', 'nan', 'None'])

    results: List[dict] = []

    # Grupp per Flextjänstnr
    for flextnr, g in df.groupby('Flextjänstnr'):
        # Förväntat antal: unika numeriska värden i gruppen
        expected_vals = sorted(set([v for v in g['_antal_karl_num'].unique() if v is not None]))
        expected = expected_vals[0] if expected_vals else None
        inconsistent_expected = len(expected_vals) > 1

        # Räkna unika externa referenser
        unique_extern_refs = sorted({r for r in g['_extern_ref'].tolist() if r not in ('', 'nan', 'None')})
        actual_count = len(unique_extern_refs)

        # Kontrollera avvikelse
        is_deviation = False
        if expected is None:
            is_deviation = True
        elif actual_count != expected:
            is_deviation = True
        if inconsistent_expected:
            is_deviation = True

        if is_deviation:
            # Några identifierande fält för kontext (hämtas från första rad i gruppen)
            aff = g['Affärsenhet'].iat[0] if 'Affärsenhet' in g.columns else ""
            status = g['Status'].iat[0] if 'Status' in g.columns else ""
            flexaddr = g['Flexplatsadress'].iat[0] if 'Flexplatsadress' in g.columns else ""
            flextyp = g['Flextyp'].iat[0] if 'Flextyp' in g.columns else ""

            results.append({
                "Affärsenhet": aff,
                "Status": status,
                "Flexplatsadress": flexaddr,
                "Flextjänstnr": flextnr,
                "Flextyp": flextyp,
                "Antal på flextjänsten": expected,
                "Antal aktiva individer": actual_count,
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
@bp.route('/upload/individer_check', methods=['POST'])
def individer_check_upload():
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
    output_filename = f"avvikelser_individer_{session_id}.xlsx"
    output_path = UPLOAD_FOLDER / output_filename

    file.save(input_path)

    try:
        deviations = process_karl(input_path, output_path)
    except ValueError as e:
        flash(str(e))
        return redirect(url_for('index'))
    except Exception:
        flash('Fel vid bearbetning av filen')
        return redirect(url_for('index'))

    if deviations > 0:
        message = f"{deviations} flextjänster har avvikande antalsvärde mot antalet aktiva individer"
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
