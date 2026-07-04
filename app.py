"""
Interface web locale pour generer les releves de notes (et attestations de
reussite) a partir d'un PV de deliberation. Parcours : choix de la filiere
("/") -> depot des fichiers ("/upload") -> generation ("/generate"). Lancer
avec : python app.py, puis ouvrir http://127.0.0.1:5000 dans un navigateur.
"""
import os
import shutil
import traceback
import uuid
from datetime import datetime
from pathlib import Path

from flask import Flask, flash, redirect, render_template, request, send_from_directory, url_for
from werkzeug.utils import secure_filename

import generate_releves as gen

BASE_DIR = Path(__file__).resolve().parent
# Dossier contenant les modeles Exemple*.xlsx / AR*.docx : configurable via
# TEMPLATES_DIR (utilise par Docker, ou ces fichiers sont montes en volume
# plutot que copies dans l'image). Par defaut, memes fichiers qu'avant : a
# cote de app.py.
TEMPLATES_DIR = Path(os.environ.get("TEMPLATES_DIR", str(BASE_DIR)))
try:
    TEMPLATE_RELEVE_PATH = next(TEMPLATES_DIR.glob("Exemple*.xlsx"))
except StopIteration:
    raise SystemExit(
        f"Fichier modele introuvable : placez un fichier 'Exemple ....xlsx' dans {TEMPLATES_DIR}"
    )
TEMPLATE_AR_PATH = next(TEMPLATES_DIR.glob("AR*.docx"), None)
RUNS_DIR = BASE_DIR / "web_runs"
RUNS_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "releve-generator-local")
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 Mo

IN_DOCKER = os.environ.get("IN_DOCKER") == "1"


@app.context_processor
def inject_in_docker():
    return {"in_docker": IN_DOCKER}


def is_valid_batch_id(batch_id):
    return bool(batch_id) and all(c.isalnum() or c in "-_" for c in batch_id) and ".." not in batch_id


def _niveaux_valides(filiere):
    cycle = gen.cycle_de_filiere(filiere)
    return [code for code, _ in gen.NIVEAUX_PAR_CYCLE.get(cycle, [])]


@app.route("/")
def index():
    return render_template(
        "index.html",
        filieres=list(gen.FILIERES.keys()),
        niveaux_par_cycle=gen.NIVEAUX_PAR_CYCLE,
        cycle_de_filiere={label: gen.cycle_de_filiere(label) for label in gen.FILIERES},
    )


@app.route("/upload")
def upload():
    filiere = request.args.get("filiere", "")
    niveau = request.args.get("niveau", "")
    if filiere not in gen.FILIERES or niveau not in _niveaux_valides(filiere):
        flash("Merci de choisir une filière et un niveau.", "error")
        return redirect(url_for("index"))
    return render_template("upload.html", filiere=filiere, niveau=niveau)


@app.route("/generate", methods=["POST"])
def generate():
    filiere = request.form.get("filiere", "")
    niveau = request.form.get("niveau", "")
    pv_file = request.files.get("pv_file")
    coords_file = request.files.get("coords_file")
    make_pdf = request.form.get("make_pdf") == "on"

    if filiere not in gen.FILIERES or niveau not in _niveaux_valides(filiere):
        flash("Merci de choisir une filière et un niveau.", "error")
        return redirect(url_for("index"))
    if not pv_file or not pv_file.filename:
        flash("Il faut fournir le PV de délibération.", "error")
        return redirect(url_for("upload", filiere=filiere, niveau=niveau))
    if not pv_file.filename.lower().endswith(".xlsx"):
        flash(f"'{pv_file.filename}' n'est pas un fichier .xlsx valide.", "error")
        return redirect(url_for("upload", filiere=filiere, niveau=niveau))
    if coords_file and coords_file.filename and not coords_file.filename.lower().endswith(".xlsx"):
        flash(f"'{coords_file.filename}' n'est pas un fichier .xlsx valide.", "error")
        return redirect(url_for("upload", filiere=filiere, niveau=niveau))

    batch_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
    batch_dir = RUNS_DIR / batch_id
    batch_dir.mkdir(parents=True)

    pv_path = batch_dir / secure_filename(pv_file.filename)
    pv_file.save(pv_path)

    coords_path = None
    if coords_file and coords_file.filename:
        coords_path = batch_dir / secure_filename(coords_file.filename)
        coords_file.save(coords_path)

    output_dir = batch_dir / "releves"
    try:
        generated, warnings = gen.generate_all(
            [pv_path], TEMPLATE_RELEVE_PATH, output_dir,
            coords_path=coords_path,
            filiere_code=gen.FILIERES[filiere],
            ar_template_path=TEMPLATE_AR_PATH,
            niveau=niveau,
        )
    except Exception as exc:
        traceback.print_exc()
        flash(f"Erreur pendant la generation : {exc}", "error")
        shutil.rmtree(batch_dir, ignore_errors=True)
        return redirect(url_for("upload", filiere=filiere, niveau=niveau))

    pdf_by_xlsx = {}
    if make_pdf and generated:
        try:
            import export_pdf
            pdfs = export_pdf.convert_to_pdf(generated)
            pdf_by_xlsx = {x.name: p.name for x, p in zip(generated, pdfs)}
        except Exception as exc:
            traceback.print_exc()
            flash(f"Les .xlsx ont ete generes mais l'export PDF a echoue : {exc}", "error")

    ar_by_xlsx = {}
    if TEMPLATE_AR_PATH:
        for x in generated:
            ar_name = "AR_" + x.name[len("Releve_"):].rsplit(".", 1)[0] + ".docx"
            if (output_dir / ar_name).exists():
                ar_by_xlsx[x.name] = ar_name

    zip_path = batch_dir / "releves.zip"
    shutil.make_archive(str(zip_path.with_suffix("")), "zip", str(output_dir))

    students = sorted(p.name for p in generated)
    return render_template(
        "result.html",
        batch_id=batch_id,
        students=students,
        pdf_by_xlsx=pdf_by_xlsx,
        ar_by_xlsx=ar_by_xlsx,
        warnings=warnings,
        zip_name=zip_path.name,
    )


@app.route("/download/<batch_id>/<filename>")
def download(batch_id, filename):
    if not is_valid_batch_id(batch_id):
        return "Identifiant invalide", 400
    directory = RUNS_DIR / batch_id if filename == "releves.zip" else RUNS_DIR / batch_id / "releves"
    return send_from_directory(directory, filename, as_attachment=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
