from pathlib import Path
from typing import List
import pandas as pd

from flask import Blueprint, request, flash, redirect, url_for, session
from utils.file_utils import allowed_file, create_session_paths, cleanup_folder, UPLOAD_FOLDER

bp = Blueprint('hamtfrekvens_mat_rest', __name__)

def process_hamtfrekvens(input_path: Path, output_path: Path) -> int:

    df = pd.read_excel(input_path)

    # Kontrollera obligatoriska kolumner
    required_cols = {
        'Affärsenhet',
        'Kundnummer',
        'Flexplats',
        'Flexplatsadress',
        'Fraktion',
        'Hämtfrekvens',
        'Flextjänst'
    }
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Saknar kolumner: {', '.join(missing)}")

    # Normalisera fraktion och filtrera på Matavfall/Restavfall
    df['Fraktion'] = df['Fraktion'].fillna('')
    df['Fraktion_norm'] = df['Fraktion'].replace({'Restavfall nollvision': 'Restavfall'})
    df = df[df['Fraktion_norm'].isin(['Matavfall', 'Restavfall'])].copy()

    # Mappa text till numeriskt värde (hämtningar per vecka)
    freq_map = {
        'varannan vecka': 0.5,
        '1 gång i veckan': 1.0,
        '2 gånger i veckan': 2.0,
        '3 gånger i veckan': 3.0,
    }

    def to_numeric_freq(s: str):
        if pd.isna(s):
            return None
        key = str(s).strip().lower()
        return freq_map.get(key)

    results: List[dict] = []

    # Gruppera per Flexplats och jämför max-frekvenser
    for flexplats, group in df.groupby('Flexplats'):
        mats = group[group['Fraktion_norm'] == 'Matavfall'].copy()
        rests = group[group['Fraktion_norm'] == 'Restavfall'].copy()

        # Hoppa över om någon fraktion saknas
        if mats.empty or rests.empty:
            continue

        # Konvertera frekvenser till numeriska värden, filtrera bort okända
        mats['freq_num'] = mats['Hämtfrekvens'].apply(to_numeric_freq)
        rests['freq_num'] = rests['Hämtfrekvens'].apply(to_numeric_freq)

        mat_vals = sorted([v for v in mats['freq_num'].unique() if v is not None])
        rest_vals = sorted([v for v in rests['freq_num'].unique() if v is not None])

        # Om vi inte har numeriska värden, hoppa över
        if not mat_vals or not rest_vals:
            continue

        # Jämför: skriv ut endast om mat har högre max-frekvens än rest
        if max(mat_vals) > max(rest_vals):
            mat_tjanster = sorted(set(mats['Flextjänst'].astype(str)))
            rest_tjanster = sorted(set(rests['Flextjänst'].astype(str)))

            results.append({
                "Affärsenhet": group['Affärsenhet'].iat[0],
                "Kundnummer": group['Kundnummer'].iat[0],
                "Flexplats": flexplats,
                "Flexplatsadress": group['Flexplatsadress'].iat[0],
                "Matavfall hämtfrekvenser": sorted(set(mats['Hämtfrekvens'].astype(str))),
                "Restavfall hämtfrekvenser": sorted(set(rests['Hämtfrekvens'].astype(str))),

                "Matavfall flextjänster": mat_tjanster,
                "Restavfall flextjänster": rest_tjanster
            })

    out_df = pd.DataFrame(results)

    # Skriv resultat till Excel
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Avvikelser")
        workbook = writer.book
        worksheet = writer.sheets["Avvikelser"]

        header_fmt = workbook.add_format({"align": "left", "bold": True})
        for col, value in enumerate(out_df.columns):
            worksheet.write(0, col, value, header_fmt)

        cell_fmt = workbook.add_format({"align": "left"})
        worksheet.set_column(0, len(out_df.columns)-1, 25, cell_fmt)

    return len(out_df)


# Endpoint för filuppladdning och bearbetning
@bp.route('/upload/hamtfrekvens', methods=['POST'])
def hamtfrekvens_mat_rest():
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
    output_filename = f"avvikelser_hamtfrekvens_{session_id}.xlsx"
    output_path = UPLOAD_FOLDER / output_filename

    file.save(input_path)

    try:
        deviations = process_hamtfrekvens(input_path, output_path)
    except ValueError as e:
        flash(str(e))
        return redirect(url_for('index'))
    except Exception:
        flash('Fel vid bearbetning av filen')
        return redirect(url_for('index'))

    if deviations > 0:
        message = (
            f"{deviations} flexplatser har avvikelser där matavfallet har tätare hämtning än restavfallet och behöver åtgärd"
        )
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
