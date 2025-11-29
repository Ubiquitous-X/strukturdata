import os
from pathlib import Path
from flask import Flask, flash, request, session, url_for, redirect, render_template, send_from_directory, abort
from werkzeug.exceptions import RequestEntityTooLarge
from dotenv import load_dotenv
from utils.file_utils import UPLOAD_FOLDER
from views.hamtfrekvens_mat_rest import bp as hamtfrekvens_mat_rest_bp
from views.hamtfrekvens_prisdel import bp as hamtfrekvens_prisdel_bp
from views.antalsvarde_individer import bp as antalsvarde_individer_bp
from views.debiteringsgrupp_check import bp as debiteringsgrupp_check_bp

load_dotenv(".env")

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY') or 'dev-secret'
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB

app.register_blueprint(hamtfrekvens_mat_rest_bp)
app.register_blueprint(hamtfrekvens_prisdel_bp)
app.register_blueprint(antalsvarde_individer_bp)
app.register_blueprint(debiteringsgrupp_check_bp)


# Endpoint för startsidan
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


# Endpoint för success-sida
@app.route('/success', methods=['GET'])
def success():
    session_id = request.args.get('file')
    if not session_id:
        flash("Ingen session angiven.")
        return redirect(url_for('index'))

    session_key = f"result_{session_id}"
    result = session.pop(session_key, None)

    if not result:
        flash("Resultat ej hittat (sessionen kan ha gått ut).")
        return redirect(url_for('index'))

    return render_template(
        'success.html',
        deviations=result.get('deviations', 0),
        output_filename=result.get('output_filename'),
        back_endpoint=result.get('back', 'index'),
        message=result.get('message', 'Resultat')
    )


# Endpoint för nedladdning av fil
@app.route('/download/<path:filename>', methods=['GET'])
def download_file(filename):
    requested = (Path(UPLOAD_FOLDER) / filename).resolve()
    upload_resolved = Path(UPLOAD_FOLDER).resolve()

    if upload_resolved not in requested.parents and requested != upload_resolved:
        abort(404)

    if not requested.exists() or not requested.is_file():
        abort(404)

    return send_from_directory(str(upload_resolved), filename, as_attachment=True)


# Felhantering för för stora filer
@app.errorhandler(RequestEntityTooLarge)
def handle_large_file(e):
    flash("Filen är för stor. Max tillåten storlek är 10 MB")
    return redirect(url_for('index'))


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
