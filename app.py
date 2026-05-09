import os
import re
import io
import json
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

from PyPDF2 import PdfReader
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


# ============================================================
# CONFIGURACIÓN BASE
# ============================================================

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.config["SECRET_KEY"] = "cambia-esta-clave-en-produccion"
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(BASE_DIR, "instance", "database.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["UPLOAD_FOLDER"] = os.path.join(BASE_DIR, "uploads")
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024

ALLOWED_EXTENSIONS = {".pdf"}

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(os.path.join(BASE_DIR, "instance"), exist_ok=True)

db = SQLAlchemy(app)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"


# ============================================================
# MODELOS DE BASE DE DATOS
# ============================================================

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    cvs = db.relationship("CV", backref="owner", lazy=True)


class CV(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    original_name = db.Column(db.String(255), nullable=False)
    extracted_text = db.Column(db.Text, default="")
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)


class Analysis(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    job_description = db.Column(db.Text, nullable=False)
    best_candidate = db.Column(db.String(255))
    best_score = db.Column(db.Float, default=0)
    summary = db.Column(db.Text)
    user_id = db.Column(db.Integer, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# ============================================================
# MOTOR INTELIGENTE LOCAL
# ============================================================

STOPWORDS = {
    "y", "o", "de", "del", "la", "el", "los", "las", "un", "una", "unos", "unas",
    "que", "con", "para", "por", "en", "al", "lo", "su", "sus", "como", "se",
    "es", "son", "ser", "tener", "tenga", "tiene", "requiere", "requiero",
    "busco", "necesito", "candidato", "persona", "perfil", "buen", "buena",
    "muy", "tambien", "también", "este", "esta", "estos", "estas"
}

SYNONYMS = {
    "redes": ["networking", "network", "lan", "wan", "tcp/ip", "router", "switch", "cisco", "routing", "switching"],
    "telecomunicaciones": ["telecom", "fibra", "fibra optica", "fibra óptica", "redes", "conectividad"],
    "soporte": ["helpdesk", "mesa de ayuda", "soporte tecnico", "soporte técnico", "asistencia tecnica"],
    "administracion": ["administración", "gestion", "gestión", "management", "operacion", "operación"],
    "gestion": ["gestión", "administracion", "administración", "management", "coordinacion", "coordinación"],
    "ingles": ["inglés", "english", "bilingue", "bilingüe", "b2", "c1", "c2"],
    "frances": ["francés", "french"],
    "espanol": ["español", "spanish", "castellano"],
    "certificacion": ["certificación", "certificado", "diploma", "titulo", "título", "licencia"],
    "seguridad": ["security", "ciberseguridad", "cybersecurity", "firewall", "hardening"],
    "cloud": ["aws", "azure", "gcp", "nube", "cloud computing"],
    "programacion": ["programación", "python", "javascript", "developer", "desarrollador", "software"],
    "liderazgo": ["lider", "líder", "coordinador", "jefatura", "supervisor"],
    "proactivo": ["iniciativa", "autonomo", "autónomo", "mejora continua"],
    "responsable": ["comprometido", "puntual", "ordenado", "organizado"],
    "presion": ["presión", "bajo presion", "bajo presión", "trabajo bajo presión"],
    "ventas": ["comercial", "vendedor", "sales", "ejecutivo comercial"],
    "contabilidad": ["contador", "finanzas", "auditoria", "auditoría"],
    "rrhh": ["recursos humanos", "reclutamiento", "seleccion", "selección"],
}


def normalize_text(text):
    if not text:
        return ""
    text = text.lower()
    replacements = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "ñ": "n"
    }
    for a, b in replacements.items():
        text = text.replace(a, b)
    text = re.sub(r"[^a-z0-9áéíóúñü\s/+#.-]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_keywords(job_description):
    text = normalize_text(job_description)
    words = text.split()

    keywords = []
    for word in words:
        if len(word) > 2 and word not in STOPWORDS:
            keywords.append(word)

    phrases = []
    phrase_patterns = [
        "bajo presion", "bajo presión", "fibra optica", "fibra óptica",
        "recursos humanos", "gestion de proyectos", "gestión de proyectos",
        "soporte tecnico", "soporte técnico", "mesa de ayuda",
        "cisco ccna", "cisco ccnp", "seguridad informatica",
        "seguridad informática"
    ]

    for phrase in phrase_patterns:
        if normalize_text(phrase) in text:
            phrases.append(normalize_text(phrase))

    expanded = set(keywords + phrases)

    for key in list(expanded):
        if key in SYNONYMS:
            for item in SYNONYMS[key]:
                expanded.add(normalize_text(item))

    return sorted(expanded)


def extract_pdf_text(path):
    text = ""
    try:
        reader = PdfReader(path)
        for page in reader.pages:
            text += " " + (page.extract_text() or "")
    except Exception:
        text = ""
    return text.strip()


def detect_experience_years(text):
    """
    Detecta años de experiencia SOLO cuando el texto habla explícitamente de experiencia laboral.
    Evita confundir edad, fechas, teléfonos, RUT o años sueltos con experiencia.
    """
    text_norm = normalize_text(text)

    patterns = [
        r"(\d{1,2})\s*(anos|año|años|years)\s*(de)?\s*(experiencia|experiencia laboral|experiencia profesional)",
        r"(experiencia|experiencia laboral|experiencia profesional)\s*(de)?\s*(\d{1,2})\s*(anos|año|años|years)",
        r"mas de\s*(\d{1,2})\s*(anos|año|años|years)\s*(de)?\s*(experiencia|experiencia laboral|experiencia profesional)",
    ]

    detected = []

    for pattern in patterns:
        matches = re.findall(pattern, text_norm)
        for match in matches:
            for item in match:
                if str(item).isdigit():
                    value = int(item)
                    if 0 < value <= 40:
                        detected.append(value)

    return max(detected) if detected else 0


def detect_languages(text):
    text = normalize_text(text)
    langs = []
    if any(x in text for x in ["ingles", "english", "bilingue", "b2", "c1", "c2"]):
        langs.append("Inglés")
    if any(x in text for x in ["frances", "french"]):
        langs.append("Francés")
    if any(x in text for x in ["espanol", "castellano", "spanish"]):
        langs.append("Español")
    return langs


def calculate_candidate_score(job_description, cv_text):
    cv_norm = normalize_text(cv_text)
    keywords = extract_keywords(job_description)

    exact_matches = []
    fuzzy_matches = []

    for keyword in keywords:
        if keyword in cv_norm:
            exact_matches.append(keyword)
        else:
            ratio = fuzz.partial_ratio(keyword, cv_norm)
            if ratio >= 88:
                fuzzy_matches.append(keyword)

    total_keywords = max(len(keywords), 1)

    exact_score = len(exact_matches) / total_keywords
    fuzzy_score = len(fuzzy_matches) / total_keywords * 0.65

    experience_years = detect_experience_years(cv_text)
    experience_bonus = min(experience_years / 10, 1) * 0.10

    language_bonus = min(len(detect_languages(cv_text)) * 0.03, 0.09)

    final_score = (exact_score * 0.75) + fuzzy_score + experience_bonus + language_bonus
    final_score = min(final_score * 100, 100)

    matches = exact_matches + fuzzy_matches

    return round(final_score, 1), matches[:20], {
        "keywords": keywords,
        "experience_years": experience_years,
        "languages": detect_languages(cv_text)
    }


def generate_ai_style_summary(candidate_name, score, matches, metadata):
    if not matches:
        return (
            f"El perfil {candidate_name} aparece como el mejor dentro de los CVs disponibles, "
            f"pero no se detectaron coincidencias fuertes y explícitas con la descripción ingresada. "
            f"Se recomienda revisar manualmente el documento antes de avanzar."
        )

    top_matches = ", ".join(matches[:8])
    years = metadata.get("experience_years", 0)
    langs = metadata.get("languages", [])

    summary = (
        f"El perfil {candidate_name} destaca frente a los demás porque presenta coincidencias relevantes "
        f"con los criterios solicitados. Las principales señales detectadas fueron: {top_matches}. "
    )

    if years > 0:
        summary += (
            f"Además, el CV menciona explícitamente alrededor de {years} año(s) de experiencia laboral, "
            f"lo que refuerza su compatibilidad con el cargo. "
        )
    else:
        summary += (
            "No se detectó una frase explícita que indique años totales de experiencia laboral, "
            "por lo tanto este dato debe validarse manualmente durante la revisión o entrevista. "
        )

    if langs:
        summary += f"También se identificaron idiomas relevantes: {', '.join(langs)}. "

    if score >= 75:
        summary += "En conclusión, es un candidato altamente alineado para avanzar a una siguiente etapa."
    elif score >= 50:
        summary += "En conclusión, es un candidato interesante, pero requiere validación humana antes de avanzar."
    else:
        summary += "En conclusión, aunque lidera el ranking actual, su compatibilidad todavía es parcial."

    return summary

# ============================================================
# UTILIDADES
# ============================================================

def allowed_file(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


def user_upload_folder():
    folder = os.path.join(app.config["UPLOAD_FOLDER"], str(current_user.id))
    os.makedirs(folder, exist_ok=True)
    return folder


# ============================================================
# RUTAS
# ============================================================

@app.route("/")
def home():
    return render_template("home.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        if not username or not password:
            flash("Debes ingresar usuario y contraseña.", "error")
            return redirect(url_for("register"))

        existing = User.query.filter_by(username=username).first()
        if existing:
            flash("Ese usuario ya existe.", "error")
            return redirect(url_for("register"))

        user = User(
            username=username,
            password_hash=generate_password_hash(password)
        )

        db.session.add(user)
        db.session.commit()

        login_user(user)
        return redirect(url_for("dashboard"))

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        user = User.query.filter_by(username=username).first()

        if user and check_password_hash(user.password_hash, password):
            login_user(user)
            return redirect(url_for("dashboard"))

        flash("Usuario o contraseña incorrectos.", "error")
        return redirect(url_for("login"))

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("home"))


@app.route("/dashboard", methods=["GET", "POST"])
@login_required
def dashboard():
    cvs = CV.query.filter_by(user_id=current_user.id).order_by(CV.uploaded_at.desc()).all()

    results = []
    recommended = None
    ai_summary = None
    job_description = ""

    if request.method == "POST":
        job_description = request.form.get("job_description", "").strip()

        for cv in cvs:
            score, matches, metadata = calculate_candidate_score(job_description, cv.extracted_text)
            results.append({
                "id": cv.id,
                "name": cv.original_name,
                "score": score,
                "matches": matches,
                "metadata": metadata
            })

        results.sort(key=lambda item: item["score"], reverse=True)

        if results:
            recommended = results[0]
            ai_summary = generate_ai_style_summary(
                recommended["name"],
                recommended["score"],
                recommended["matches"],
                recommended["metadata"]
            )

            analysis = Analysis(
                job_description=job_description,
                best_candidate=recommended["name"],
                best_score=recommended["score"],
                summary=ai_summary,
                user_id=current_user.id
            )
            db.session.add(analysis)
            db.session.commit()

    total_cvs = len(cvs)
    total_analyses = Analysis.query.filter_by(user_id=current_user.id).count()

    return render_template(
        "dashboard.html",
        cvs=cvs,
        results=results,
        recommended=recommended,
        ai_summary=ai_summary,
        job_description=job_description,
        total_cvs=total_cvs,
        total_analyses=total_analyses
    )


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    files = request.files.getlist("cvs")

    for file in files:
        if not file or file.filename == "":
            continue

        if not allowed_file(file.filename):
            continue

        original_name = file.filename
        safe_name = secure_filename(file.filename)

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        filename = f"{timestamp}_{safe_name}"

        path = os.path.join(user_upload_folder(), filename)
        file.save(path)

        extracted_text = extract_pdf_text(path)

        cv = CV(
            filename=filename,
            original_name=original_name,
            extracted_text=extracted_text,
            user_id=current_user.id
        )

        db.session.add(cv)

    db.session.commit()
    return redirect(url_for("dashboard"))


@app.route("/download/<int:cv_id>")
@login_required
def download(cv_id):
    cv = CV.query.get_or_404(cv_id)

    if cv.user_id != current_user.id:
        return redirect(url_for("dashboard"))

    return send_from_directory(user_upload_folder(), cv.filename, as_attachment=True, download_name=cv.original_name)


@app.route("/delete/<int:cv_id>")
@login_required
def delete_cv(cv_id):
    cv = CV.query.get_or_404(cv_id)

    if cv.user_id == current_user.id:
        path = os.path.join(user_upload_folder(), cv.filename)
        if os.path.exists(path):
            os.remove(path)

        db.session.delete(cv)
        db.session.commit()

    return redirect(url_for("dashboard"))


@app.route("/delete_all")
@login_required
def delete_all():
    cvs = CV.query.filter_by(user_id=current_user.id).all()

    for cv in cvs:
        path = os.path.join(user_upload_folder(), cv.filename)
        if os.path.exists(path):
            os.remove(path)
        db.session.delete(cv)

    db.session.commit()
    return redirect(url_for("dashboard"))


@app.route("/export_excel")
@login_required
def export_excel():
    cvs = CV.query.filter_by(user_id=current_user.id).all()
    last_analysis = Analysis.query.filter_by(user_id=current_user.id).order_by(Analysis.created_at.desc()).first()

    job_description = last_analysis.job_description if last_analysis else ""

    rows = []
    for cv in cvs:
        score, matches, metadata = calculate_candidate_score(job_description, cv.extracted_text)
        rows.append({
            "candidate": cv.original_name,
            "score": score,
            "matches": ", ".join(matches[:12]),
            "experience": metadata.get("experience_years", 0),
            "languages": ", ".join(metadata.get("languages", [])),
            "uploaded_at": cv.uploaded_at.strftime("%Y-%m-%d %H:%M")
        })

    rows.sort(key=lambda item: item["score"], reverse=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Ranking ATS"

    headers = ["Ranking", "Candidato", "Score", "Coincidencias", "Años Exp.", "Idiomas", "Fecha subida"]
    ws.append(headers)

    for index, row in enumerate(rows, start=1):
        ws.append([
            index,
            row["candidate"],
            row["score"],
            row["matches"],
            row["experience"],
            row["languages"],
            row["uploaded_at"]
        ])

    header_fill = PatternFill("solid", fgColor="111827")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D1D5DB")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    widths = [12, 35, 12, 55, 14, 24, 22]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = width

    if len(rows) > 0:
        table_ref = f"A1:G{len(rows) + 1}"
        table = Table(displayName="TablaRankingATS", ref=table_ref)
        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        ws.add_table(table)

    ws.freeze_panes = "A2"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="ranking_ats_profesional.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


with app.app_context():
    db.create_all()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
