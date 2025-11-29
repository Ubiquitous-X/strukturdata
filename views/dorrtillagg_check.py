from pathlib import Path
from typing import List, Optional, Dict
import pandas as pd

from flask import Blueprint, request, flash, redirect, url_for, session
from utils.file_utils import allowed_file, create_session_paths, cleanup_folder, UPLOAD_FOLDER

bp = Blueprint('dorrtillagg_check', __name__)


# Mappning från text till numeriskt värde (hämtningar per vecka)
FREQ_MAP: Dict[str, float] = {
    "varannan vecka": 0.5,
    "1 gång i veckan": 1.0,
    "2 gånger i veckan": 2.0,
    "3 gånger i veckan": 3.0,
    "var 4:e vecka": 1.0 / 4.0,
    "var 8:e vecka": 1.0 / 8.0,
}


def _norm(s: Optional[str]) -> str:
    if pd.isna(s):
        return ""
    return str(s).strip()


def _norm_lower(s: Optional[str]) -> str:
    return _norm(s).lower()


def freq_to_num(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    key = _norm_lower(s)
    return FREQ_MAP.get(key)


def process_dorrtillagg(input_path: Path, output_path: Path) -> int:

    df = pd.read_excel(input_path)

    # Kontrollera obligatoriska kolumner
    required_cols = {
        "Affärsenhet",
        "Kundnummer",
        "Flexplats",
        "Flexplatsadress",
        "Flextjänst",
        "Flexgrupp",
        "Flextyp",
        "Hämtfrekvens",
    }
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Saknar kolumner: {', '.join(missing)}")

    results: List[dict] = []

    # Grupp per Flexplats
    for flexplats, group in df.groupby("Flexplats"):
        group = group.copy()
        group["__flexgrupp_norm"] = group["Flexgrupp"].astype(str).str.strip().str.lower()

        flexgrupp_set = set(group["__flexgrupp_norm"].unique())

        # Finns endast kärl på flexplatsen ignoreras den
        if flexgrupp_set == {"kärl"}:
            continue

        # Finns endast 'tillägg' är det en avvikelse
        if flexgrupp_set == {"tillägg"}:
            for _, row in group.iterrows():
                results.append({
                    "Affärsenhet": row.get("Affärsenhet"),
                    "Kundnummer": row.get("Kundnummer"),
                    "Flexplats": row.get("Flexplats"),
                    "Flexplatsadress": row.get("Flexplatsadress"),
                    "Flextjänst": row.get("Flextjänst"),
                    "Flextyp": row.get("Flextyp"),
                    "Hämtfrekvens": row.get("Hämtfrekvens"),
                    "Orsak": "Endast flextjänst för dörrtillägg finns på flexplatsen"
                })
            continue

        # Hitta tätaste frekvens bland kärl när det finns både kärl och tillägg
        karlar = group[group["__flexgrupp_norm"] == "kärl"].copy()
        tillagg = group[group["__flexgrupp_norm"] == "tillägg"].copy()
        if tillagg.empty:
            tillagg = group[group["__flexgrupp_norm"] == "tillagg"].copy()

        # Konvertera kärlens frekvenser till numeriska värden
        karlar["__freq_num"] = karlar["Hämtfrekvens"].apply(freq_to_num)

        # Bestäm det tätaste intervallet bland kärl (högst numeriskt värde)
        max_num = karlar["__freq_num"].max()
        karlar_with_max = karlar[karlar["__freq_num"] == max_num]
        expected_text = karlar_with_max["Hämtfrekvens"].iloc[0]

        # Jämför varje tillägg
        for _, trow in tillagg.iterrows():
            t_freq_text = trow.get("Hämtfrekvens")
            t_freq_num = freq_to_num(t_freq_text)

            # Jämför mot tätaste intervallet för kärl
            if t_freq_num != max_num:
                results.append({
                    "Affärsenhet": trow.get("Affärsenhet"),
                    "Kundnummer": trow.get("Kundnummer"),
                    "Flexplats": trow.get("Flexplats"),
                    "Flexplatsadress": trow.get("Flexplatsadress"),
                    "Flextjänst": trow.get("Flextjänst"),
                    "Flextyp": trow.get("Flextyp"),
                    "Hämtfrekvens": t_freq_text,
                    "Kärlens tätaste hämtfrekvens": expected_text,
                    "Orsak": f"Dörrilläggets hämtfrekvens '{t_freq_text}' avviker från kärlens tätaste '{expected_text}'"
                })

    out_df = pd.DataFrame(results)

    # Bestäm kolumnordning
    if not out_df.empty:
        cols_order = [
            "Affärsenhet",
            "Kundnummer",
            "Flexplats",
            "Flexplatsadress",
            "Flextjänst",
            "Flextyp",
            "Hämtfrekvens",
            "Kärlens tätaste hämtfrekvens",
            "Orsak",
        ]
        cols_order = [c for c in cols_order if c in out_df.columns]
        out_df = out_df.loc[:, cols_order]


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
@bp.route('/upload/dorrtillagg_check', methods=['POST'])
def dorrtillagg_upload():
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
    output_filename = f"avvikelser_dorrtillagg_{session_id}.xlsx"
    output_path = UPLOAD_FOLDER / output_filename

    file.save(input_path)

    try:
        deviations = process_dorrtillagg(input_path, output_path)
    except ValueError as e:
        flash(str(e))
        return redirect(url_for('index'))
    except Exception:
        flash('Fel vid bearbetning av filen')
        return redirect(url_for('index'))

    if deviations > 0:
        message = f"{deviations} flexplatser har mismatch i hämtfrekvens mellan dörrtillägg/kärl och behöver åtgärd"
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
