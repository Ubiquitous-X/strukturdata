import os
import uuid
import pandas as pd
from flask import Flask, flash, request, url_for, redirect, render_template, send_from_directory
from dotenv import load_dotenv

load_dotenv('.env')
APP_ROOT = os.getcwd()
UPLOAD_FOLDER = os.path.join(os.getcwd(), "files")
ALLOWED_EXT = {'xlsx'}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY')
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # max 10 MB


def sanitize_cell(s):
    if not isinstance(s, str):
        return s
    if s and s[0] in ('=', '+', '-', '@'):
        return "'" + s
    return s


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':

        if 'file' not in request.files:
            flash('Ingen fil i anropet')
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            flash('Du måste välja en fil')
            return redirect(request.url)
        
        # Rensa gamla filer i /files
        for f in os.listdir(UPLOAD_FOLDER):
            try:
                os.remove(os.path.join(UPLOAD_FOLDER, f))
            except:
                pass

        # Kontrollera filändelse
        if '.' in file.filename and file.filename.rsplit('.', 1)[1].lower() == 'xlsx':

            # Skapa unikt ID för denna bearbetning
            session_id = uuid.uuid4().hex[:8]
            input_path = os.path.join(UPLOAD_FOLDER, f"inkommande_{session_id}.xlsx")
            output_path = os.path.join(UPLOAD_FOLDER, f"avvikelser_hämtfrekvens_{session_id}.xlsx")

            # Spara inkommande fil
            file.save(input_path)

            # Läs in Excel
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
                flash(f"Filen saknar följande kolumner: {', '.join(missing)}")
                return redirect(url_for('upload_file'))

            # Bearbeta data
            df['Fraktion'] = df['Fraktion'].fillna('')
            df['Fraktion_norm'] = df['Fraktion'].replace({'Restavfall nollvision': 'Restavfall'})
            df = df[df['Fraktion_norm'].isin(['Matavfall', 'Restavfall'])].copy()

            results = []

            for flexplats, group in df.groupby('Flexplats'):
                mats = group[group['Fraktion_norm'] == 'Matavfall']
                rests = group[group['Fraktion_norm'] == 'Restavfall']

                if mats.empty or rests.empty:
                    continue

                mat_set = set(mats['Hämtfrekvens'].str.strip().str.lower())
                rest_set = set(rests['Hämtfrekvens'].str.strip().str.lower())

                mat_tjanster = sorted(set(mats['Flextjänst'].astype(str)))
                rest_tjanster = sorted(set(rests['Flextjänst'].astype(str)))

                if mat_set != rest_set:
                    results.append({
                        "Affärsenhet": group['Affärsenhet'].iat[0],
                        "Kundnummer": group['Kundnummer'].iat[0],
                        "Flexplats": flexplats,
                        "Flexplatsadress": group['Flexplatsadress'].iat[0],
                        "Matavfall hämtfrekvenser": sorted(mat_set),
                        "Restavfall hämtfrekvenser": sorted(rest_set),
                        "Matavfall flextjänster": mat_tjanster,
                        "Restavfall flextjänster": rest_tjanster
                    })

            out_df = pd.DataFrame(results)

            # Spara Excel
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                out_df.to_excel(writer, index=False, sheet_name="Avvikelser")

                workbook = writer.book
                worksheet = writer.sheets["Avvikelser"]

                header_fmt = workbook.add_format({"align": "left", "bold": True})
                for col, value in enumerate(out_df.columns):
                    worksheet.write(0, col, value, header_fmt)

                cell_fmt = workbook.add_format({"align": "left"})
                worksheet.set_column(0, len(out_df.columns)-1, 25, cell_fmt)

            return redirect(url_for('success', deviations=len(out_df), file=session_id))

        else:
            flash('Endast Excel-filer (.xlsx) tillåtna')
            return redirect(request.url)

    return render_template('index.html')


@app.route('/success', methods=['GET', 'POST'])
def success():
    session_id = request.args.get('file')
    deviations = request.args.get('deviations', 0)
    output_filename = f"avvikelser_hämtfrekvens_{session_id}.xlsx"

    if request.method == 'POST':
        if request.form.get('download') == 'Hämta filen':
            return send_from_directory(UPLOAD_FOLDER, output_filename, as_attachment=True)
        else:
            return redirect(url_for('upload_file'))

    return render_template('success.html',
                           deviations=deviations,
                           output_filename=output_filename)


@app.route('/mall', methods=['GET'])
def example():
    return send_from_directory(APP_ROOT, 'exempelfil.xlsx', as_attachment=True)


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
