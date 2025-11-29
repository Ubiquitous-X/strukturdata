from pathlib import Path
import pandas as pd

from flask import Blueprint, request, flash, redirect, url_for, session
from utils.file_utils import allowed_file, create_session_paths, cleanup_folder, UPLOAD_FOLDER

bp = Blueprint('hamtfrekvens_prisdel', __name__)


def process_prisdel(input_path: Path, output_path: Path) -> int:

    df = pd.read_excel(input_path)

    # Kontrollera obligatoriska kolumner
    required_cols = {
        'Affärsenhet',
        'Kundnummer',
        'Avtalsnummer',
        'Flexplatsadress',
        'Flextjänst',
        'Hämtfrekvens',
        'Prisdel',
        'Status flextjänst'
    }

    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Saknar kolumner: {', '.join(missing)}")

    col_freq = 'Hämtfrekvens'
    col_pris = 'Prisdel'

    cols = list(df.columns)

    # Hjälpfunktion för kontroll och förklaring
    def check_and_reason(hamt, pris):

        hamt_s = str(hamt).strip().lower()
        pris_s = str(pris).strip().lower()
        if hamt_s in pris_s:
            return True, ""
        else:
            return False, f"Hämtfrekvens '{hamt}' finns inte i prisdelen på avtalet"

    # Kör kontrollen radvis och spara anledning för avvikelse
    check_results = df.apply(lambda r: check_and_reason(r[col_freq], r[col_pris]), axis=1)
    ok_mask = check_results.apply(lambda x: x[0])
    reasons = check_results.apply(lambda x: x[1])

    deviations_df = df.loc[~ok_mask].copy()
    deviations_df['Orsak'] = reasons.loc[~ok_mask].values

    # Välj identifierande kolumner
    ident_cols = []
    for wanted in ("Affärsenhet", "Kundnummer", "Avtalsnummer", "Flexplatsadress"):
        for c in cols:
            if str(c).strip().lower() == wanted.lower():
                ident_cols.append(c)
                break

    out_cols = ident_cols + [col_freq, col_pris, 'Orsak']
    if not ident_cols:
        # Om inga identifierare finns, inkludera de tre första kolumnerna
        out_cols = cols[:3] + [col_freq, col_pris, 'Orsak']

    out_df = deviations_df.loc[:, out_cols].copy()

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
@bp.route('/upload/prisdel_check', methods=['POST'])
def prisdel_check_upload():
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
    output_filename = f"avvikelser_prisdel_{session_id}.xlsx"
    output_path = UPLOAD_FOLDER / output_filename

    file.save(input_path)

    try:
        deviations = process_prisdel(input_path, output_path)
    except ValueError as e:
        flash(str(e))
        return redirect(url_for('index'))
    except Exception:
        flash('Fel vid bearbetning av filen')
        return redirect(url_for('index'))

    if deviations > 0:
        message = f"{deviations} flextjänster har mismatch mellan hämtfrekvensen och prisdelen på avtalet."
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
