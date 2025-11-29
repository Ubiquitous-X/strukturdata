from pathlib import Path
import uuid
from werkzeug.utils import secure_filename
from typing import Tuple

ALLOWED_EXT = {"xlsx"}

# Skapa katalog om den inte finns
BASE_DIR = Path.cwd()
UPLOAD_FOLDER = BASE_DIR / "files"
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)


def allowed_file(filename: str) -> bool:
    # Kontrollera filändelse
    if not filename:
        return False
    name = secure_filename(filename)
    return "." in name and name.rsplit(".", 1)[1].lower() in ALLOWED_EXT


def create_session_paths(filename: str) -> Tuple[Path, str]:
    # Skapa unika filnamn för uppladdning baserat på session id
    session_id = uuid.uuid4().hex[:8]
    safe = secure_filename(filename)
    input_name = f"inkommande_{session_id}_{safe}"
    return UPLOAD_FOLDER / input_name, session_id


def cleanup_folder():
    # Rensa uppladdningsmappen
    for p in UPLOAD_FOLDER.iterdir():
        try:
            p.unlink()
        except Exception:
            pass
