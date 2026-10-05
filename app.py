"""
Conciliador Bancario Web — v3
Autenticación · Usuarios · Multi-banco · Historial
Correr local: python app.py
Railway:      gunicorn app:app --timeout 120
"""

import os, re, json, base64, io, sqlite3
from pathlib import Path
from datetime import datetime, date
from functools import wraps
from flask import (Flask, request, jsonify, send_file,
                   render_template_string, session, redirect, url_for)
import anthropic, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from werkzeug.security import generate_password_hash, check_password_hash

# ── CONFIG ───────────────────────────────────────────────────────
API_KEY       = os.environ.get("ANTHROPIC_API_KEY", "sk-ant-...")
SECRET_KEY    = os.environ.get("SECRET_KEY", "cncl-lsa-2026-secret")
DB_PATH       = os.environ.get("DB_PATH", "conciliador.db")
MASTER_PASS   = os.environ.get("MASTER_PASSWORD", "admin123")
UPLOAD_FOLDER = "uploads_tmp"
UMBRAL_MATCH  = 4

# Bancos por país
BANK_CONFIG = {
    "AR": {
        "Ciudad": {
            "label": "Banco Ciudad",
            "filter_starts": ["TRANSFER"],
            "exclude_cuits": ["30708635754", "30711747709"],
        }
    },
    "PY": {
        "ITAU": {
            "label": "ITAÚ Paraguay",
            "filter_starts": ["TRANSFER"],
            "exclude_cuits": [],
        },
        "UENO": {
            "label": "UENO Paraguay",
            "filter_starts": ["TRANSFER"],
            "exclude_cuits": [],
        }
    }
}

COUNTRY_LABELS = {"AR": "🇦🇷 Argentina", "PY": "🇵🇾 Paraguay"}
# ─────────────────────────────────────────────────────────────────

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ── Base de datos ─────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            username   TEXT UNIQUE NOT NULL,
            password   TEXT NOT NULL,
            role       TEXT NOT NULL DEFAULT 'user',
            country    TEXT NOT NULL DEFAULT 'AR',
            active     INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS reconciliations (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id        INTEGER NOT NULL,
            username       TEXT NOT NULL,
            bank           TEXT NOT NULL,
            country        TEXT NOT NULL,
            excel_name     TEXT,
            total          INTEGER,
            verified       INTEGER,
            sin_comp       INTEGER,
            sin_match_cnt  INTEGER,
            results        TEXT,
            sin_match_data TEXT,
            created_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    cur = conn.execute("SELECT id FROM users WHERE role='master'")
    if not cur.fetchone():
        conn.execute(
            "INSERT INTO users (username, password, role, country) VALUES (?,?,?,?)",
            ("admin", generate_password_hash(MASTER_PASS), "master", "AR")
        )
        conn.commit()
    conn.close()


# ── Auth ──────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def master_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        if session.get("role") != "master":
            return redirect(url_for("index"))
        return f(*args, **kwargs)
    return decorated


# ── Helpers ───────────────────────────────────────────────────────

def find_col(keys, hints):
    lower = [k.lower().strip() for k in keys]
    for h in hints:
        for i, k in enumerate(lower):
            if h in k: return keys[i]
    return None

def normalize_monto(v):
    if v is None or v == "": return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip()
    negativo = s.startswith("-")
    s = re.sub(r"[^\d,\.]", "", s)
    if not s: return 0.0
    if "," in s and "." in s: result = float(s.replace(".", "").replace(",", "."))
    elif "," in s: result = float(s.replace(",", "."))
    else: result = float(s) if s else 0.0
    return -result if negativo else result

def normalize_date(v):
    if not v: return ""
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$", s)
    if m:
        y = ("20" + m.group(3)) if len(m.group(3)) == 2 else m.group(3)
        return f"{y}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
    m = re.match(r"^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$", s)
    if m: return f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
    if "T" in s: return s.split("T")[0]
    return s[:10]

def clean_cuit(c): return re.sub(r"[^\d]", "", str(c)) if c else ""

def cuit_en_desc(desc):
    m = re.search(r"(?:^|\s)(\d{10,11})[\s\-]", str(desc))
    return m.group(1) if m else ""

def file_to_b64(path):
    with open(path, "rb") as f:
        return base64.standard_b64encode(f.read()).decode()

def get_media_type(path):
    return {".pdf": "application/pdf", ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg", ".png": "image/png"}.get(Path(path).suffix.lower(), "")

def banks_for_user(country, role):
    if role == "master":
        banks = {}
        for c, bks in BANK_CONFIG.items():
            for k, v in bks.items():
                banks[k] = v["label"] + f" ({COUNTRY_LABELS.get(c,c)})"
        return banks
    return {k: v["label"] for k, v in BANK_CONFIG.get(country, {}).items()}


# ── IA ────────────────────────────────────────────────────────────

def extract_comprobante(path):
    client = anthropic.Anthropic(api_key=API_KEY)
    b64 = file_to_b64(path)
    mt  = get_media_type(path)
    is_img = Path(path).suffix.lower() in {".jpg", ".jpeg", ".png"}
    resp = client.messages.create(
        model="claude-opus-4-5", max_tokens=700,
        messages=[{"role": "user", "content": [
            {"type": "image" if is_img else "document",
             "source": {"type": "base64", "media_type": mt, "data": b64}},
            {"type": "text", "text": (
                "Comprobante de transferencia bancaria. Devolvé SOLO JSON con: "
                "fecha (DD/MM/YYYY), monto (número sin símbolo ni puntos de miles), "
                "moneda, cuit_origen (solo dígitos), nombre_origen, banco_origen, referencia. "
                "Null si no aparece. Sin backticks."
            )}
        ]}]
    )
    text = "".join(b.text for b in resp.content if hasattr(b, "text"))
    try:
        return json.loads(re.sub(r"```json|```", "", text).strip())
    except:
        return {"monto": None, "fecha": None, "error": "parse_error"}


# ── Match ─────────────────────────────────────────────────────────

def match_score(mov_fecha, mov_monto, mov_desc, comp):
    score, razones = 0, []
    c_monto = normalize_monto(comp.get("monto"))
    if c_monto and mov_monto:
        diff = abs(mov_monto - c_monto) / max(mov_monto, c_monto)
        if diff < 0.01: score += 3; razones.append("monto exacto")
        elif diff < 0.05: score += 1; razones.append("monto aprox")
    cuit_c = clean_cuit(comp.get("cuit_origen", ""))
    cuit_d = cuit_en_desc(mov_desc)
    if cuit_c and cuit_d and cuit_c[-10:] == cuit_d[-10:]:
        score += 3; razones.append("CUIT")
    c_fecha = normalize_date(comp.get("fecha", ""))
    if mov_fecha and c_fecha:
        if mov_fecha == c_fecha: score += 2; razones.append("fecha exacta")
        else:
            try:
                d1 = date.fromisoformat(mov_fecha)
                d2 = date.fromisoformat(c_fecha)
                if abs((d1 - d2).days) <= 3: score += 1; razones.append("fecha cercana")
            except: pass
    return score, razones


# ── Filtrar ingresos ──────────────────────────────────────────────

def filtrar_ingresos(df, cols, bank_key):
    # Buscar config del banco en cualquier país
    cfg = {}
    for country_banks in BANK_CONFIG.values():
        if bank_key in country_banks:
            cfg = country_banks[bank_key]
            break
    mask = df[cols["monto"]].apply(normalize_monto) > 0
    if cols["desc"]:
        starts  = cfg.get("filter_starts", ["TRANSFER"])
        desc_up = df[cols["desc"]].fillna("").str.strip().str.upper()
        mask_d  = pd.Series([False] * len(df), index=df.index)
        for s in starts:
            mask_d |= desc_up.str.startswith(s)
        mask &= mask_d
        for cuit_ex in cfg.get("exclude_cuits", []):
            mask &= ~df[cols["desc"]].fillna("").str.contains(cuit_ex, na=False)
    return df[mask].copy()


# ── Conciliar ─────────────────────────────────────────────────────

def conciliar(planilla_path, comp_paths, bank_key):
    df = pd.read_excel(planilla_path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    HINTS = {
        "fecha": ["fecha", "date", "día", "dia"],
        "monto": ["haber", "crédito", "credito", "importe", "monto", "ingreso", "amount", "entrada", "credit"],
        "desc":  ["descripcion", "descripción", "concepto", "detalle", "movimiento", "referencia", "desc", "glosa"],
    }
    cols = {k: find_col(df.columns.tolist(), v) for k, v in HINTS.items()}
    if not cols["monto"]:
        raise ValueError(f"No se encontró columna de monto. Columnas: {df.columns.tolist()}")
    ing = filtrar_ingresos(df, cols, bank_key)
    comprobantes = []
    for p in comp_paths:
        try:
            data = extract_comprobante(p)
            comprobantes.append({"file": Path(p).name, "data": data})
        except Exception as e:
            comprobantes.append({"file": Path(p).name, "data": {"monto": None, "fecha": None}, "error": str(e)})
    rows = []
    for _, mov in ing.iterrows():
        fecha = normalize_date(mov.get(cols["fecha"], "") if cols["fecha"] else "")
        monto = normalize_monto(mov.get(cols["monto"], 0))
        desc  = str(mov.get(cols["desc"], "") if cols["desc"] else "").strip()
        best_score, best_comp, best_razones = 0, None, []
        for c in comprobantes:
            s, r = match_score(fecha, monto, desc, c["data"])
            if s > best_score: best_score, best_comp, best_razones = s, c, r
        verificado = best_score >= UMBRAL_MATCH
        rows.append({
            "fecha": fecha, "descripcion": desc, "importe": round(monto, 2),
            "estado": "Verificado" if verificado else "Sin comprobante",
            "comprobante": best_comp["file"] if verificado else "",
            "match": " + ".join(best_razones) if verificado else "",
            "score": best_score,
        })
    comp_usados = {r["comprobante"] for r in rows if r["comprobante"]}
    sin_match = []
    for c in comprobantes:
        if c["file"] not in comp_usados:
            d = c.get("data", {})
            sin_match.append({
                "file":   c["file"],
                "fecha":  normalize_date(d.get("fecha", "")),
                "monto":  normalize_monto(d.get("monto") or 0),
                "nombre": d.get("nombre_origen") or "—",
                "banco":  d.get("banco_origen") or "—",
                "cuit":   d.get("cuit_origen") or "—",
            })
    return {"rows": rows, "sin_match": sin_match}


# ── Exportadores ──────────────────────────────────────────────────

def export_excel(rows, sin_match=None, bank_label="", username=""):
    wb = Workbook()
    ws = wb.active
    ws.title = "Conciliacion"
    verde    = PatternFill("solid", fgColor="C6EFCE")
    amarillo = PatternFill("solid", fgColor="FFEB9C")
    hfill    = PatternFill("solid", fgColor="1F4E79")
    hfont    = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    bfont    = Font(name="Arial", size=10)
    thin     = Side(style="thin", color="BFBFBF")
    border   = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.append(["Fecha", "Descripción", "Importe", "Estado", "Comprobante", "Criterios"])
    for cell in ws[1]:
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal="center"); cell.border = border
    for r in rows:
        ws.append([r["fecha"], r["descripcion"], r["importe"], r["estado"], r["comprobante"], r["match"]])
        fill = verde if r["estado"] == "Verificado" else amarillo
        for cell in ws[ws.max_row]:
            cell.fill = fill; cell.font = bfont; cell.border = border
            cell.alignment = Alignment(horizontal="left")
    for col in ws.columns:
        w = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(w + 4, 55)
    ws.freeze_panes = "A2"
    if sin_match:
        ws2 = wb.create_sheet("Sin match")
        ws2.append(["Archivo", "Fecha", "Monto", "Remitente", "Banco", "CUIT"])
        for cell in ws2[1]:
            cell.fill = hfill; cell.font = hfont
            cell.alignment = Alignment(horizontal="center"); cell.border = border
        rojo = PatternFill("solid", fgColor="FFCCCC")
        for r in sin_match:
            ws2.append([r["file"], r["fecha"], r["monto"], r["nombre"], r["banco"], r["cuit"]])
            for cell in ws2[ws2.max_row]:
                cell.fill = rojo; cell.font = bfont; cell.border = border
        for col in ws2.columns:
            w = max(len(str(c.value or "")) for c in col)
            ws2.column_dimensions[get_column_letter(col[0].column)].width = min(w + 4, 50)
        ws2.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf

def export_pdf(rows, sin_match=None, bank_label="", username=""):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=2*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    t_style = ParagraphStyle("t", parent=styles["Heading1"], fontSize=14,
                             textColor=colors.HexColor("#1F4E79"), spaceAfter=6)
    s_style = ParagraphStyle("s", parent=styles["Normal"], fontSize=9,
                             textColor=colors.grey, spaceAfter=14)
    c_style = ParagraphStyle("c", parent=styles["Normal"], fontSize=8, leading=11)
    total = len(rows); verif = sum(1 for r in rows if r["estado"] == "Verificado")
    story = [
        Paragraph(f"Conciliación — {bank_label}", t_style),
        Paragraph(f"Usuario: {username}  ·  {datetime.now().strftime('%d/%m/%Y %H:%M')}  ·  "
                  f"Total: {total}  ·  Verificados: {verif}  ·  Sin comprobante: {total-verif}", s_style),
    ]
    tdata = [["Fecha", "Descripción", "Importe", "Estado", "Comprobante", "Match"]]
    for r in rows:
        tdata.append([r["fecha"], Paragraph(r["descripcion"][:60], c_style),
                      f"${r['importe']:,.0f}", r["estado"],
                      Paragraph(r["comprobante"], c_style), r["match"]])
    t = Table(tdata, colWidths=[2.2*cm,8*cm,2.8*cm,3.2*cm,4.5*cm,5.5*cm], repeatRows=1)
    cmds = [
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E79")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("FONTNAME",(0,1),(-1,-1),"Helvetica"),
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),5),
    ]
    for i, r in enumerate(rows, 1):
        fill = colors.HexColor("#C6EFCE") if r["estado"]=="Verificado" else colors.HexColor("#FFEB9C")
        cmds.append(("BACKGROUND",(0,i),(-1,i),fill))
    t.setStyle(TableStyle(cmds))
    story.append(t)
    if sin_match:
        story.append(Spacer(1,1*cm))
        story.append(Paragraph("Comprobantes sin coincidencia", t_style))
        t2d = [["Archivo","Fecha","Monto","Remitente","Banco"]]
        for r in sin_match:
            t2d.append([r["file"],r["fecha"]or"—",
                        f"${r['monto']:,.0f}"if r["monto"]else"—",
                        Paragraph(r["nombre"]or"—",c_style),r["banco"]or"—"])
        t2 = Table(t2d,colWidths=[4*cm,2.5*cm,3*cm,8*cm,4*cm],repeatRows=1)
        t2.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#C0392B")),
            ("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),8),
            ("FONTNAME",(0,1),(-1,-1),"Helvetica"),
            ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#FFCCCC")),
            ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ]))
        story.append(t2)
    doc.build(story)
    buf.seek(0)
    return buf


# ── HTML ──────────────────────────────────────────────────────────

LOGIN_HTML = """
<!DOCTYPE html><html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Conciliador — Ingresar</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>✅</text></svg>">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#F0F4FF;display:flex;align-items:center;justify-content:center;min-height:100vh}
.card{background:#fff;border-radius:12px;padding:2.5rem;width:360px;box-shadow:0 4px 24px rgba(0,0,0,.1)}
.logo{text-align:center;font-size:48px;margin-bottom:1rem}
h1{text-align:center;font-size:20px;color:#1F4E79;margin-bottom:.25rem}
.sub{text-align:center;font-size:13px;color:#6B7280;margin-bottom:2rem}
label{display:block;font-size:13px;font-weight:600;color:#374151;margin-bottom:4px}
input{width:100%;padding:.65rem .85rem;border:1px solid #D1D5DB;border-radius:8px;font-size:14px;margin-bottom:1rem;outline:none}
input:focus{border-color:#1F4E79;box-shadow:0 0 0 3px rgba(31,78,121,.1)}
.btn{width:100%;padding:.75rem;background:#1F4E79;color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer}
.btn:hover{background:#163A5F}
.error{background:#FEF2F2;color:#991B1B;border:1px solid #FCA5A5;border-radius:8px;padding:.65rem 1rem;font-size:13px;margin-bottom:1rem}
</style></head><body>
<div class="card">
  <div class="logo">✅</div>
  <h1>Conciliador Bancario</h1>
  <p class="sub">La Santaniana S.A.</p>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="POST">
    <label>Usuario</label>
    <input type="text" name="username" autocomplete="username" required autofocus>
    <label>Contraseña</label>
    <input type="password" name="password" autocomplete="current-password" required>
    <button class="btn" type="submit">Ingresar</button>
  </form>
</div>
</body></html>
"""

ADMIN_HTML = """
<!DOCTYPE html><html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Conciliador — Admin</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>✅</text></svg>">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#F7F8FA;color:#1A1A2E;font-size:14px}
header{background:#1F4E79;color:#fff;padding:.85rem 2rem;display:flex;align-items:center;justify-content:space-between}
header h1{font-size:17px;font-weight:600}
.header-right{display:flex;align-items:center;gap:1rem;font-size:13px}
.header-right a{color:#fff;text-decoration:none;opacity:.8}
.header-right a:hover{opacity:1}
.container{max-width:900px;margin:2rem auto;padding:0 1.5rem}
.card{background:#fff;border:1px solid #E2E4E8;border-radius:10px;padding:1.5rem;margin-bottom:1.25rem}
.card-title{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#6B7280;margin-bottom:1rem}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;padding:8px 10px;font-size:11px;font-weight:600;color:#6B7280;border-bottom:2px solid #E2E4E8}
td{padding:8px 10px;border-bottom:1px solid #E2E4E8;vertical-align:middle}
tr:last-child td{border-bottom:none}
.badge{font-size:11px;font-weight:600;padding:2px 8px;border-radius:4px}
.badge-master{background:#EAF3DE;color:#3B6D11}
.badge-user{background:#E6F1FB;color:#185FA5}
.badge-ar{background:#EAF3DE;color:#3B6D11}
.badge-py{background:#FEF3C7;color:#92400E}
.badge-off{background:#F3F4F6;color:#6B7280}
.btn-sm{font-size:12px;padding:4px 10px;border:1px solid #E2E4E8;border-radius:6px;cursor:pointer;background:#fff}
.btn-danger{border-color:#FCA5A5;color:#991B1B}
.btn-success{border-color:#86EFAC;color:#166534}
.form-grid{display:grid;grid-template-columns:1fr 1fr 1fr 1fr auto;gap:.75rem;align-items:end}
label{display:block;font-size:12px;font-weight:600;color:#374151;margin-bottom:4px}
input,select{width:100%;padding:.55rem .75rem;border:1px solid #D1D5DB;border-radius:7px;font-size:13px;outline:none}
input:focus,select:focus{border-color:#1F4E79}
.btn-primary{padding:.6rem 1.25rem;background:#1F4E79;color:#fff;border:none;border-radius:7px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap}
.btn-primary:hover{background:#163A5F}
.msg{padding:.65rem 1rem;border-radius:8px;font-size:13px;margin-bottom:1rem}
.msg-ok{background:#F0FDF4;color:#166534;border:1px solid #86EFAC}
.msg-err{background:#FEF2F2;color:#991B1B;border:1px solid #FCA5A5}
</style></head><body>
<header>
  <h1>⚙ Gestión de Usuarios</h1>
  <div class="header-right">
    <span>{{ session_username }}</span>
    <a href="/">← Volver</a>
    <a href="/logout">Salir</a>
  </div>
</header>
<div class="container">
  {% if msg %}<div class="msg msg-ok">{{ msg }}</div>{% endif %}
  {% if err %}<div class="msg msg-err">{{ err }}</div>{% endif %}

  <div class="card">
    <div class="card-title">Crear usuario</div>
    <form method="POST" action="/admin/crear">
      <div class="form-grid">
        <div><label>Usuario</label><input type="text" name="username" required></div>
        <div><label>Contraseña</label><input type="password" name="password" required></div>
        <div><label>País</label>
          <select name="country">
            <option value="AR">🇦🇷 Argentina</option>
            <option value="PY">🇵🇾 Paraguay</option>
          </select>
        </div>
        <div><label>Rol</label>
          <select name="role">
            <option value="user">Usuario</option>
            <option value="master">Master</option>
          </select>
        </div>
        <div><label>&nbsp;</label><button class="btn-primary" type="submit">Crear</button></div>
      </div>
    </form>
  </div>

  <div class="card">
    <div class="card-title">Usuarios ({{ users|length }})</div>
    <table>
      <thead><tr>
        <th>Usuario</th><th>Rol</th><th>País</th><th>Estado</th><th>Creado</th><th>Acciones</th>
      </tr></thead>
      <tbody>
      {% for u in users %}
      <tr>
        <td><strong>{{ u.username }}</strong></td>
        <td><span class="badge badge-{{ u.role }}">{{ u.role }}</span></td>
        <td><span class="badge badge-{{ u.country|lower }}">{{ '🇦🇷 AR' if u.country=='AR' else '🇵🇾 PY' }}</span></td>
        <td><span class="badge {{ 'badge-user' if u.active else 'badge-off' }}">{{ 'Activo' if u.active else 'Inactivo' }}</span></td>
        <td style="color:#6B7280;font-size:12px">{{ u.created_at[:10] }}</td>
        <td>
          {% if u.role != 'master' %}
          <form method="POST" action="/admin/toggle/{{ u.id }}" style="display:inline">
            <button class="btn-sm {{ 'btn-danger' if u.active else 'btn-success' }}" type="submit">
              {{ 'Desactivar' if u.active else 'Activar' }}
            </button>
          </form>
          {% endif %}
        </td>
      </tr>
      {% endfor %}
      </tbody>
    </table>
  </div>
</div>
</body></html>
"""

APP_HTML = """
<!DOCTYPE html><html lang="es"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Conciliador Bancario</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>✅</text></svg>">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#F7F8FA;color:#1A1A2E;font-size:14px}
header{background:#1F4E79;color:#fff;padding:.85rem 2rem;display:flex;align-items:center;justify-content:space-between}
header h1{font-size:17px;font-weight:600}
.header-right{display:flex;align-items:center;gap:1rem;font-size:13px}
.country-badge{background:rgba(255,255,255,.15);padding:3px 10px;border-radius:20px;font-size:12px}
.header-right a{color:#fff;text-decoration:none;opacity:.8;font-size:13px}
.header-right a:hover{opacity:1}
.tabs-bar{background:#fff;border-bottom:1px solid #E2E4E8;padding:0 2rem;display:flex;gap:.25rem}
.tab{padding:.75rem 1.25rem;font-size:13px;font-weight:600;color:#6B7280;border:none;background:none;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px}
.tab.active{color:#1F4E79;border-bottom-color:#1F4E79}
.tab-content{display:none}.tab-content.active{display:block}
.container{max-width:1000px;margin:1.5rem auto;padding:0 1.5rem}
.card{background:#fff;border:1px solid #E2E4E8;border-radius:10px;padding:1.5rem;margin-bottom:1.25rem}
.card-title{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#6B7280;margin-bottom:1rem}
.upload-grid{display:grid;grid-template-columns:1fr 1fr;gap:1rem}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem}
.drop{border:1.5px dashed #D1D5DB;border-radius:8px;padding:1.5rem;text-align:center;cursor:pointer;transition:.15s}
.drop:hover{background:#F0F4FF;border-color:#93C5FD}
.file-list{margin-top:.75rem;display:flex;flex-direction:column;gap:6px}
.file-item{display:flex;align-items:center;gap:8px;font-size:12px;padding:5px 10px;background:#F9FAFB;border-radius:6px;border:1px solid #E2E4E8}
.badge{font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px}
.badge-xlsx{background:#EAF3DE;color:#3B6D11}
.badge-pdf{background:#FAECE7;color:#993C1D}
.badge-img{background:#E6F1FB;color:#185FA5}
.fname{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.rm{cursor:pointer;color:#9CA3AF;font-size:15px;line-height:1}
.rm:hover{color:#E24B4A}
select.bank-sel{width:100%;padding:.65rem .85rem;border:1px solid #D1D5DB;border-radius:8px;font-size:14px;outline:none}
select.bank-sel:focus{border-color:#1F4E79}
.run-btn{width:100%;padding:.85rem;font-size:15px;font-weight:600;background:#1F4E79;color:#fff;border:none;border-radius:8px;cursor:pointer;transition:.15s}
.run-btn:hover:not(:disabled){background:#163A5F}
.run-btn:disabled{opacity:.45;cursor:not-allowed}
.progress-bar{height:5px;background:#E2E4E8;border-radius:3px;overflow:hidden;margin-top:.75rem}
.progress-fill{height:100%;background:#1D9E75;border-radius:3px;transition:width .4s}
.status-msg{font-size:13px;color:#6B7280;text-align:center;margin-top:.5rem;min-height:18px}
.hidden{display:none!important}
.stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:1.25rem}
.stat{background:#F9FAFB;border:1px solid #E2E4E8;border-radius:8px;padding:.85rem 1rem}
.stat-label{font-size:11px;color:#6B7280;margin-bottom:4px}
.stat-value{font-size:22px;font-weight:700}
.green{color:#1D9E75}.amber{color:#854F0B}
.results-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:.85rem}
.export-row{display:flex;gap:8px}
.btn-export{padding:.45rem 1rem;font-size:13px;font-weight:500;border:1px solid #E2E4E8;background:#fff;border-radius:6px;cursor:pointer}
.btn-export:hover{background:#F0F4FF}
.btn-xlsx{border-color:#3B6D11;color:#3B6D11}
.btn-pdf{border-color:#993C1D;color:#993C1D}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;padding:8px 10px;font-size:11px;font-weight:600;color:#6B7280;border-bottom:2px solid #E2E4E8;white-space:nowrap}
td{padding:8px 10px;border-bottom:1px solid #E2E4E8;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr.verified td{background:#C6EFCE}
tr.missing td{background:#FFEB9C}
tr.sinmatch td{background:#FFCCCC}
.chip{display:inline-flex;align-items:center;gap:4px;font-size:11px;font-weight:600;padding:3px 9px;border-radius:4px}
.chip-v{background:rgba(0,0,0,.08);color:#166534}
.chip-m{background:rgba(0,0,0,.08);color:#854F0B}
.chip-r{background:rgba(0,0,0,.08);color:#991B1B}
.dot{width:6px;height:6px;border-radius:50%}
.dot-g{background:#1D9E75}.dot-a{background:#EF9F27}.dot-r{background:#E24B4A}
.comp-name{font-size:11px;color:#185FA5;margin-top:2px}
.match-tag{font-size:11px;color:#6B7280}
.sinmatch-title{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#991B1B;margin-bottom:.75rem}
.error-box{background:#FEF2F2;border:1px solid #FCA5A5;border-radius:8px;padding:.85rem 1rem;margin-top:.75rem;font-size:13px;color:#991B1B}
.scroll-wrap{overflow-x:auto}
/* Historial */
.hist-badge{font-size:11px;font-weight:600;padding:2px 8px;border-radius:4px}
.hist-ar{background:#EAF3DE;color:#3B6D11}
.hist-py{background:#FEF3C7;color:#92400E}
.hist-btn{font-size:11px;padding:3px 9px;border:1px solid #E2E4E8;border-radius:5px;cursor:pointer;background:#fff;margin-right:4px}
.hist-btn:hover{background:#F0F4FF}
@media(max-width:700px){.upload-grid,.form-row,.stats-grid{grid-template-columns:1fr}}
</style></head><body>

<header>
  <h1>✅ Conciliador Bancario</h1>
  <div class="header-right">
    <span class="country-badge">{{ country_label }}</span>
    <span>{{ username }}</span>
    {% if role == 'master' %}<a href="/admin">⚙ Admin</a>{% endif %}
    <a href="/logout">Salir</a>
  </div>
</header>

<div class="tabs-bar">
  <button class="tab active" onclick="showTab('conciliar',this)">Nueva Conciliación</button>
  <button class="tab" onclick="showTab('historial',this);loadHistorial()">Historial</button>
</div>

<div class="container">

  <!-- TAB: CONCILIAR -->
  <div id="tab-conciliar" class="tab-content active">
    <div class="card">
      <div class="card-title">Configuración</div>
      <div class="form-row">
        <div>
          <div style="font-size:12px;font-weight:600;color:#374151;margin-bottom:4px">Banco</div>
          <select class="bank-sel" id="bank-sel">
            {% for key, label in banks.items() %}
            <option value="{{ key }}">{{ label }}</option>
            {% endfor %}
          </select>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Archivos</div>
      <div class="upload-grid">
        <div>
          <div class="drop" id="drop-xlsx" onclick="document.getElementById('inp-xlsx').click()">
            <div style="font-size:24px;margin-bottom:6px">📊</div>
            <div style="font-weight:600">Planilla de movimientos</div>
            <div style="font-size:13px;color:#6B7280;margin-top:6px">Excel (.xlsx) del banco</div>
          </div>
          <input type="file" id="inp-xlsx" accept=".xlsx,.xls,.csv" style="display:none">
          <div class="file-list" id="list-xlsx"></div>
        </div>
        <div>
          <div class="drop" id="drop-comp" onclick="document.getElementById('inp-comp').click()">
            <div style="font-size:24px;margin-bottom:6px">🗂</div>
            <div style="font-weight:600">Comprobantes de transferencia</div>
            <div style="font-size:13px;color:#6B7280;margin-top:6px">PDF, JPG o PNG — varios a la vez</div>
          </div>
          <input type="file" id="inp-comp" accept=".pdf,.jpg,.jpeg,.png" multiple style="display:none">
          <div class="file-list" id="list-comp"></div>
        </div>
      </div>
    </div>

    <button class="run-btn" id="run-btn" disabled onclick="runConciliation()">Conciliar transferencias</button>
    <div class="progress-bar hidden" id="prog-wrap"><div class="progress-fill" id="prog-fill" style="width:0%"></div></div>
    <div class="status-msg" id="status-msg"></div>
    <div id="error-area"></div>

    <div id="results-section" class="hidden" style="margin-top:1.5rem">
      <div class="stats-grid" id="stats-grid"></div>
      <div class="card">
        <div class="results-header">
          <div class="card-title" style="margin-bottom:0">Transferencias</div>
          <div class="export-row">
            <button class="btn-export btn-xlsx" onclick="exportFile('xlsx')">⬇ Excel</button>
            <button class="btn-export btn-pdf"  onclick="exportFile('pdf')">⬇ PDF</button>
          </div>
        </div>
        <div class="scroll-wrap">
          <table><thead><tr>
            <th style="width:90px">Fecha</th>
            <th>Descripción</th>
            <th style="width:120px;text-align:right">Importe</th>
            <th style="width:145px">Estado</th>
            <th>Comprobante / Criterios</th>
          </tr></thead><tbody id="results-body"></tbody></table>
        </div>
      </div>

      <div id="sinmatch-section" class="hidden">
        <div class="card">
          <div class="results-header">
            <div class="sinmatch-title">⚠ Comprobantes sin coincidencia</div>
            <div style="font-size:12px;color:#6B7280" id="sinmatch-count"></div>
          </div>
          <div class="scroll-wrap">
            <table><thead><tr>
              <th>Archivo</th><th style="width:90px">Fecha</th>
              <th style="width:110px;text-align:right">Monto</th>
              <th>Remitente</th><th>Banco</th><th>CUIT</th>
            </tr></thead><tbody id="sinmatch-body"></tbody></table>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- TAB: HISTORIAL -->
  <div id="tab-historial" class="tab-content">
    <div class="card">
      <div class="results-header">
        <div class="card-title" style="margin-bottom:0">Historial de conciliaciones</div>
        <div style="font-size:12px;color:#6B7280" id="hist-count"></div>
      </div>
      <div class="scroll-wrap">
        <table><thead><tr>
          <th style="width:130px">Fecha y hora</th>
          <th style="width:120px">Banco</th>
          {% if role == 'master' %}<th>Usuario</th>{% endif %}
          <th style="width:80px;text-align:center">Total</th>
          <th style="width:90px;text-align:center">Verificados</th>
          <th style="width:110px;text-align:center">Sin comprobante</th>
          <th>Archivo</th>
          <th style="width:130px">Exportar</th>
        </tr></thead><tbody id="hist-body"><tr><td colspan="8" style="text-align:center;color:#9CA3AF;padding:2rem">Cargando...</td></tr></tbody></table>
      </div>
    </div>
  </div>

</div>

<script>
var xlsxFile=null, compFiles=[], lastResults=[], lastSinMatch=[], isMaster={{ 'true' if role=='master' else 'false' }};

function ext(n){return(n||'').split('.').pop().toLowerCase();}
function badgeClass(n){var e=ext(n);if(['xlsx','xls','csv'].includes(e))return'badge-xlsx';return e==='pdf'?'badge-pdf':'badge-img';}
function badgeLabel(n){var e=ext(n);if(e==='xlsx'||e==='xls')return'XLSX';if(e==='csv')return'CSV';if(e==='pdf')return'PDF';return'IMG';}

function renderXlsx(){
  document.getElementById('list-xlsx').innerHTML=xlsxFile
    ?'<div class="file-item"><span class="badge '+badgeClass(xlsxFile.name)+'">'+badgeLabel(xlsxFile.name)+'</span><span class="fname">'+xlsxFile.name+'</span><span class="rm" onclick="xlsxFile=null;renderXlsx();check()">×</span></div>':'';
}
function renderComp(){
  document.getElementById('list-comp').innerHTML=compFiles.map(function(f,i){
    return '<div class="file-item"><span class="badge '+badgeClass(f.name)+'">'+badgeLabel(f.name)+'</span><span class="fname">'+f.name+'</span><span class="rm" onclick="compFiles.splice('+i+',1);renderComp();check()">×</span></div>';
  }).join('');
}
function check(){document.getElementById('run-btn').disabled=!(xlsxFile&&compFiles.length>0);}

document.getElementById('inp-xlsx').onchange=function(e){if(e.target.files[0]){xlsxFile=e.target.files[0];renderXlsx();check();}};
document.getElementById('inp-comp').onchange=function(e){compFiles=compFiles.concat(Array.from(e.target.files));renderComp();check();};

['drop-xlsx','drop-comp'].forEach(function(id){
  var el=document.getElementById(id);
  el.addEventListener('dragover',function(e){e.preventDefault();el.style.background='#F0F4FF';});
  el.addEventListener('dragleave',function(){el.style.background='';});
  el.addEventListener('drop',function(e){
    e.preventDefault();el.style.background='';
    var files=Array.from(e.dataTransfer.files);
    if(id==='drop-xlsx'){var f=files.find(function(f){return['xlsx','xls','csv'].includes(ext(f.name));});if(f){xlsxFile=f;renderXlsx();check();}}
    else{compFiles=compFiles.concat(files.filter(function(f){return['pdf','jpg','jpeg','png'].includes(ext(f.name));}));renderComp();check();}
  });
});

function showTab(name,btn){
  document.querySelectorAll('.tab-content').forEach(function(t){t.classList.remove('active');});
  document.querySelectorAll('.tab').forEach(function(t){t.classList.remove('active');});
  document.getElementById('tab-'+name).classList.add('active');
  if(btn) btn.classList.add('active');
}

function setProgress(pct,msg){
  document.getElementById('prog-fill').style.width=pct+'%';
  document.getElementById('status-msg').textContent=msg;
}

async function runConciliation(){
  document.getElementById('error-area').innerHTML='';
  document.getElementById('results-section').classList.add('hidden');
  document.getElementById('prog-wrap').classList.remove('hidden');
  document.getElementById('run-btn').disabled=true;
  setProgress(5,'Subiendo archivos...');
  var fd=new FormData();
  fd.append('planilla',xlsxFile);
  compFiles.forEach(function(f){fd.append('comprobantes',f);});
  fd.append('bank',document.getElementById('bank-sel').value);
  try{
    setProgress(15,'Procesando con IA — puede tardar 1-2 minutos...');
    var resp=await fetch('/conciliar',{method:'POST',body:fd});
    var data=await resp.json();
    if(data.error){showError(data.error);return;}
    setProgress(100,'Conciliación completada.');
    lastResults=data.rows; lastSinMatch=data.sin_match||[];
    renderResults(data.rows); renderSinMatch(data.sin_match||[]);
  }catch(e){showError('Error de conexión: '+e.message);}
  finally{document.getElementById('run-btn').disabled=false;}
}

function showError(msg){
  document.getElementById('error-area').innerHTML='<div class="error-box">'+msg+'</div>';
  document.getElementById('prog-wrap').classList.add('hidden');
  document.getElementById('run-btn').disabled=false;
}

function renderResults(rows){
  var total=rows.length,verif=rows.filter(function(r){return r.estado==='Verificado';}).length;
  var totalM=rows.reduce(function(a,r){return a+r.importe;},0);
  document.getElementById('stats-grid').innerHTML=
    '<div class="stat"><div class="stat-label">Transferencias</div><div class="stat-value">'+total+'</div></div>'+
    '<div class="stat"><div class="stat-label">Verificados</div><div class="stat-value green">'+verif+'</div></div>'+
    '<div class="stat"><div class="stat-label">Sin comprobante</div><div class="stat-value amber">'+(total-verif)+'</div></div>'+
    '<div class="stat"><div class="stat-label">Total importe</div><div class="stat-value" style="font-size:15px">$'+totalM.toLocaleString('es-AR',{minimumFractionDigits:0})+'</div></div>';
  document.getElementById('results-body').innerHTML=rows.map(function(r){
    return '<tr class="'+(r.estado==='Verificado'?'verified':'missing')+'">'+
      '<td style="color:#6B7280">'+r.fecha+'</td>'+
      '<td>'+r.descripcion+'</td>'+
      '<td style="text-align:right;font-weight:600">$'+r.importe.toLocaleString('es-AR',{minimumFractionDigits:2})+'</td>'+
      '<td>'+(r.estado==='Verificado'
        ?'<div class="chip chip-v"><div class="dot dot-g"></div>Verificado</div>'
        :'<div class="chip chip-m"><div class="dot dot-a"></div>Sin comprobante</div>')+'</td>'+
      '<td>'+(r.comprobante?'<div class="comp-name">'+r.comprobante+'</div>':'')+(r.match?'<div class="match-tag">'+r.match+'</div>':'—')+'</td>'+
      '</tr>';
  }).join('');
  document.getElementById('results-section').classList.remove('hidden');
}

function renderSinMatch(items){
  var s=document.getElementById('sinmatch-section');
  if(!items||items.length===0){s.classList.add('hidden');return;}
  document.getElementById('sinmatch-count').textContent=items.length+' comprobante(s)';
  document.getElementById('sinmatch-body').innerHTML=items.map(function(r){
    return '<tr class="sinmatch">'+
      '<td><span class="chip chip-r">'+r.file+'</span></td>'+
      '<td style="color:#6B7280">'+r.fecha+'</td>'+
      '<td style="text-align:right;font-weight:600">'+(r.monto?'$'+r.monto.toLocaleString('es-AR',{minimumFractionDigits:0}):'—')+'</td>'+
      '<td>'+r.nombre+'</td><td>'+r.banco+'</td>'+
      '<td style="font-family:monospace;font-size:11px">'+r.cuit+'</td></tr>';
  }).join('');
  s.classList.remove('hidden');
}

function exportFile(fmt){
  if(!lastResults.length)return;
  var fd=new FormData();
  fd.append('rows',JSON.stringify(lastResults));
  fd.append('sin_match',JSON.stringify(lastSinMatch));
  fd.append('format',fmt);
  fd.append('bank_label',document.getElementById('bank-sel').options[document.getElementById('bank-sel').selectedIndex].text);
  fetch('/exportar',{method:'POST',body:fd}).then(function(r){return r.blob();}).then(function(blob){
    var url=URL.createObjectURL(blob);
    var a=document.createElement('a');a.href=url;
    a.download=fmt==='xlsx'?'conciliacion.xlsx':'conciliacion.pdf';
    a.click();URL.revokeObjectURL(url);
  });
}

var histLoaded=false;
function loadHistorial(){
  if(histLoaded)return;
  histLoaded=true;
  fetch('/historial').then(function(r){return r.json();}).then(function(data){
    document.getElementById('hist-count').textContent=data.length+' registros';
    if(!data.length){
      document.getElementById('hist-body').innerHTML='<tr><td colspan="8" style="text-align:center;color:#9CA3AF;padding:2rem">No hay conciliaciones registradas todavía.</td></tr>';
      return;
    }
    document.getElementById('hist-body').innerHTML=data.map(function(h){
      var countryClass=h.country==='PY'?'hist-py':'hist-ar';
      var bankLabel=h.bank+(h.country?' ('+h.country+')':'');
      return '<tr>'+
        '<td style="font-size:12px;color:#6B7280">'+h.created_at.substring(0,16).replace('T',' ')+'</td>'+
        '<td><span class="hist-badge '+countryClass+'">'+h.bank+'</span></td>'+
        (isMaster?'<td style="font-size:12px">'+h.username+'</td>':'')+
        '<td style="text-align:center">'+h.total+'</td>'+
        '<td style="text-align:center;color:#1D9E75;font-weight:600">'+h.verified+'</td>'+
        '<td style="text-align:center;color:#854F0B">'+h.sin_comp+'</td>'+
        '<td style="font-size:12px;color:#6B7280">'+h.excel_name+'</td>'+
        '<td>'+
          '<button class="hist-btn" onclick="histExport('+h.id+',\'xlsx\')">⬇ Excel</button>'+
          '<button class="hist-btn" onclick="histExport('+h.id+',\'pdf\')">⬇ PDF</button>'+
        '</td></tr>';
    }).join('');
  });
}

function histExport(id,fmt){
  fetch('/historial/exportar',{method:'POST',body:JSON.stringify({id:id,format:fmt}),headers:{'Content-Type':'application/json'}})
    .then(function(r){return r.blob();}).then(function(blob){
      var url=URL.createObjectURL(blob);
      var a=document.createElement('a');a.href=url;
      a.download=fmt==='xlsx'?'conciliacion_'+id+'.xlsx':'conciliacion_'+id+'.pdf';
      a.click();URL.revokeObjectURL(url);
    });
}
</script>
</body></html>
"""


# ── Rutas ─────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        user = conn.execute(
            "SELECT * FROM users WHERE username=? AND active=1", (username,)
        ).fetchone()
        conn.close()
        if user and check_password_hash(user["password"], password):
            session["user_id"]  = user["id"]
            session["username"] = user["username"]
            session["role"]     = user["role"]
            session["country"]  = user["country"]
            return redirect(url_for("index"))
        error = "Usuario o contraseña incorrectos."
    return render_template_string(LOGIN_HTML, error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    banks = banks_for_user(session["country"], session["role"])
    country_label = COUNTRY_LABELS.get(session["country"], session["country"])
    return render_template_string(APP_HTML,
        username=session["username"],
        role=session["role"],
        country_label=country_label,
        banks=banks,
    )

@app.route("/conciliar", methods=["POST"])
@login_required
def conciliar_route():
    if API_KEY.startswith("sk-ant-..."):
        return jsonify({"error": "API key no configurada."})
    try:
        planilla  = request.files["planilla"]
        comps     = request.files.getlist("comprobantes")
        bank_key  = request.form.get("bank", "Ciudad")
        plan_path = os.path.join(UPLOAD_FOLDER, "planilla" + Path(planilla.filename).suffix)
        planilla.save(plan_path)
        comp_paths = []
        for f in comps:
            p = os.path.join(UPLOAD_FOLDER, f.filename)
            f.save(p); comp_paths.append(p)
        result = conciliar(plan_path, comp_paths, bank_key)
        rows      = result["rows"]
        sin_match = result["sin_match"]
        # Guardar en historial
        verif    = sum(1 for r in rows if r["estado"] == "Verificado")
        sin_comp = len(rows) - verif
        # Determinar country del banco
        country = session["country"]
        conn = get_db()
        conn.execute(
            "INSERT INTO reconciliations (user_id,username,bank,country,excel_name,total,verified,sin_comp,sin_match_cnt,results,sin_match_data) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (session["user_id"], session["username"], bank_key, country,
             planilla.filename, len(rows), verif, sin_comp, len(sin_match),
             json.dumps(rows), json.dumps(sin_match))
        )
        conn.commit(); conn.close()
        return jsonify({"rows": rows, "sin_match": sin_match})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/exportar", methods=["POST"])
@login_required
def exportar_route():
    rows      = json.loads(request.form["rows"])
    sin_match = json.loads(request.form.get("sin_match", "[]"))
    fmt        = request.form["format"]
    bank_label = request.form.get("bank_label", "")
    username   = session["username"]
    if fmt == "xlsx":
        buf = export_excel(rows, sin_match, bank_label, username)
        return send_file(buf, download_name="conciliacion.xlsx", as_attachment=True,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        buf = export_pdf(rows, sin_match, bank_label, username)
        return send_file(buf, download_name="conciliacion.pdf", as_attachment=True,
                         mimetype="application/pdf")

@app.route("/historial")
@login_required
def historial_route():
    conn = get_db()
    if session["role"] == "master":
        rows = conn.execute(
            "SELECT id,username,bank,country,excel_name,total,verified,sin_comp,sin_match_cnt,created_at FROM reconciliations ORDER BY created_at DESC"
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT id,username,bank,country,excel_name,total,verified,sin_comp,sin_match_cnt,created_at FROM reconciliations WHERE user_id=? ORDER BY created_at DESC",
            (session["user_id"],)
        ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/historial/exportar", methods=["POST"])
@login_required
def historial_exportar():
    data       = request.get_json()
    rec_id     = data["id"]
    fmt        = data["format"]
    conn       = get_db()
    rec        = conn.execute("SELECT * FROM reconciliations WHERE id=?", (rec_id,)).fetchone()
    conn.close()
    if not rec:
        return jsonify({"error": "Registro no encontrado"}), 404
    if session["role"] != "master" and rec["user_id"] != session["user_id"]:
        return jsonify({"error": "Sin permisos"}), 403
    rows      = json.loads(rec["results"])
    sin_match = json.loads(rec["sin_match_data"] or "[]")
    bank_label = rec["bank"]
    username   = rec["username"]
    if fmt == "xlsx":
        buf = export_excel(rows, sin_match, bank_label, username)
        return send_file(buf, download_name=f"conciliacion_{rec_id}.xlsx", as_attachment=True,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        buf = export_pdf(rows, sin_match, bank_label, username)
        return send_file(buf, download_name=f"conciliacion_{rec_id}.pdf", as_attachment=True,
                         mimetype="application/pdf")

@app.route("/admin")
@master_required
def admin():
    conn  = get_db()
    users = conn.execute("SELECT * FROM users ORDER BY created_at DESC").fetchall()
    conn.close()
    return render_template_string(ADMIN_HTML,
        users=users,
        session_username=session["username"],
        msg=request.args.get("msg"),
        err=request.args.get("err"),
    )

@app.route("/admin/crear", methods=["POST"])
@master_required
def admin_crear():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    country  = request.form.get("country", "AR")
    role     = request.form.get("role", "user")
    if not username or not password:
        return redirect(url_for("admin") + "?err=Completá todos los campos")
    try:
        conn = get_db()
        conn.execute(
            "INSERT INTO users (username,password,role,country) VALUES (?,?,?,?)",
            (username, generate_password_hash(password), role, country)
        )
        conn.commit(); conn.close()
        return redirect(url_for("admin") + "?msg=Usuario creado correctamente")
    except sqlite3.IntegrityError:
        return redirect(url_for("admin") + "?err=El usuario ya existe")

@app.route("/admin/toggle/<int:uid>", methods=["POST"])
@master_required
def admin_toggle(uid):
    conn = get_db()
    user = conn.execute("SELECT active,role FROM users WHERE id=?", (uid,)).fetchone()
    if user and user["role"] != "master":
        conn.execute("UPDATE users SET active=? WHERE id=?", (0 if user["active"] else 1, uid))
        conn.commit()
    conn.close()
    return redirect(url_for("admin") + "?msg=Usuario actualizado")


# ── Main ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", 5000))
    host = "0.0.0.0" if os.environ.get("RAILWAY_ENVIRONMENT") else "127.0.0.1"
    if host == "127.0.0.1":
        print("=" * 50)
        print("  Conciliador Bancario v3")
        print("  http://localhost:5000")
        print("  Usuario master: admin / admin123")
        print("=" * 50)
    app.run(debug=False, host=host, port=port)

init_db()
