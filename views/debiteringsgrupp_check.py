from pathlib import Path
from typing import List
import pandas as pd

from flask import Blueprint, request, flash, redirect, url_for, session
from utils.file_utils import allowed_file, create_session_paths, cleanup_folder, UPLOAD_FOLDER

bp = Blueprint('debiteringsgrupp_check', __name__)


def normalize(s):
    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def process_debiteringsgrupp(input_path: Path, output_path: Path) -> int:
    """
    Kontrollerar debiteringsgrupp enligt regler:
      - Ignorera rader där Debiteringsgrupp är i ignored_set.
      - För SEVAB (affärsenhet som börjar med 'sevab'): Debiteringsgrupp måste vara 'Månad'.
      - För EEM (affärsenhet som börjar med 'eem'): beroende på Prislista ska debiteringsgrupp vara:
          - 'ÅVM Fritidshus' -> 'Månad maj-sept'
          - 'ÅVM En- och två bostadshus' -> 'Månad'
    """
    df = pd.read_excel(input_path)

    # Kontrollera obligatoriska kolumner
    required_cols = {
        'Affärsenhet',
        'Kundnummer',
        'Avtalsnummer',
        'Debiteringsgrupp',
        'Prislista',
        'Avtalsstatus'
    }
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Saknar kolumner: {', '.join(missing)}")

    ignored_set = {normalize(x) for x in ("Varannan månad", "BRI", "Kvartal")}
    eem_map = {
        normalize("ÅVM Fritidshus"): "Månad maj-sept",
        normalize("ÅVM En- och två bostadshus"): "Månad",
    }

    results: List[dict] = []

    for i, row in df.iterrows():
        aff = row.get("Affärsenhet")
        deb_group = row.get("Debiteringsgrupp")
        prislista = row.get("Prislista")

        deb_norm = normalize(deb_group)
        aff_norm = normalize(aff)
        pris_norm = normalize(prislista)

        # Ignorera om debiteringsgruppen är i ignored_set
        if deb_norm in ignored_set:
            continue

        is_dev = False
        reason = ""

        # Matcha affärsenhet via startswith (t.ex. "EEM Återvinning")
        if aff_norm.startswith("sevab"):
            expected = "månad"
            if deb_norm != expected:
                is_dev = True
                reason = f"Affärsenhet SEVAB förväntar Debiteringsgrupp 'Månad', hittade '{deb_group}'"
        elif aff_norm.startswith("eem"):
            if pris_norm in eem_map:
                expected_full = eem_map[pris_norm]
                if normalize(expected_full) != deb_norm:
                    is_dev = True
                    reason = (
                        f"Affärsenhet EEM med Prislista '{prislista}' förväntar "
                        f"Debiteringsgrupp '{expected_full}', hittade '{deb_group}'"
                    )

        if is_dev:
            results.append({
                "Affärsenhet": aff,
                "Kundnummer": row.get("Kundnummer"),
                "Avtalsnummer": row.get("Avtalsnummer"),
                "Debiteringsgrupp": deb_group,
                "Prislista": prislista,
                "Avtalsstatus": row.get("Avtalsstatus"),
                "Orsak": reason
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
@bp.route('/upload/debiteringsgrupp_check', methods=['POST'])
def debiteringsgrupp_upload():
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
    output_filename = f"avvikelser_debiteringsgrupp_{session_id}.xlsx"
    output_path = UPLOAD_FOLDER / output_filename

    file.save(input_path)

    try:
        deviations = process_debiteringsgrupp(input_path, output_path)
    except ValueError as e:
        flash(str(e))
        return redirect(url_for('index'))
    except Exception:
        flash('Fel vid bearbetning av filen')
        return redirect(url_for('index'))

    if deviations > 0:
        message = f"{deviations} avtal ligger på felaktig debiteringsgrupp och behöver åtgärd"
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
