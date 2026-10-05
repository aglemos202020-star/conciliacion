"""
Conciliador Bancario Web
========================
Correr local: python app.py  →  http://localhost:5000
Railway:      gunicorn app:app --timeout 120
"""

import os, re, json, base64, io
from pathlib import Path
from datetime import datetime, date
from flask import Flask, request, jsonify, send_file, render_template_string
import anthropic, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ── CONFIG ───────────────────────────────────────────────────────
API_KEY       = os.environ.get("ANTHROPIC_API_KEY", "sk-ant-...")
UPLOAD_FOLDER = "uploads_tmp"
UMBRAL_MATCH  = 4

BANK_CONFIG = {
    "Ciudad": {
        "label": "Banco Ciudad (Argentina)",
        "filter_starts": ["TRANSFER"],
        "exclude_cuits": ["30708635754", "30711747709"],
    },
    "ITAU": {
        "label": "ITAÚ (Paraguay)",
        "filter_starts": ["TRANSFER"],
        "exclude_cuits": [],
    },
    "UENO": {
        "label": "UENO (Paraguay)",
        "filter_starts": ["TRANSFER"],
        "exclude_cuits": [],
    },
}
# ─────────────────────────────────────────────────────────────────

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


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
    neg = s.startswith("-")
    s = re.sub(r"[^\d,\.]", "", s)
    if not s: return 0.0
    if "," in s and "." in s: r = float(s.replace(".", "").replace(",", "."))
    elif "," in s: r = float(s.replace(",", "."))
    else: r = float(s) if s else 0.0
    return -r if neg else r

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
    return {".pdf":"application/pdf",".jpg":"image/jpeg",
            ".jpeg":"image/jpeg",".png":"image/png"}.get(Path(path).suffix.lower(),"")


# ── IA ────────────────────────────────────────────────────────────

def extract_comprobante(path):
    client = anthropic.Anthropic(api_key=API_KEY)
    b64 = file_to_b64(path)
    mt  = get_media_type(path)
    is_img = Path(path).suffix.lower() in {".jpg",".jpeg",".png"}
    resp = client.messages.create(
        model="claude-opus-4-5", max_tokens=700,
        messages=[{"role":"user","content":[
            {"type":"image" if is_img else "document",
             "source":{"type":"base64","media_type":mt,"data":b64}},
            {"type":"text","text":(
                "Comprobante de transferencia bancaria. Devolvé SOLO JSON con: "
                "fecha (DD/MM/YYYY), monto (numero sin simbolo ni puntos de miles), "
                "moneda, cuit_origen (solo digitos), nombre_origen, banco_origen, referencia. "
                "Null si no aparece. Sin backticks."
            )}
        ]}]
    )
    text = "".join(b.text for b in resp.content if hasattr(b,"text"))
    try: return json.loads(re.sub(r"```json|```","",text).strip())
    except: return {"monto":None,"fecha":None,"error":"parse_error"}


# ── Match ─────────────────────────────────────────────────────────

def match_score(mov_fecha, mov_monto, mov_desc, comp):
    score, razones = 0, []
    c_monto = normalize_monto(comp.get("monto"))
    if c_monto and mov_monto:
        diff = abs(mov_monto - c_monto) / max(mov_monto, c_monto)
        if diff < 0.01: score += 3; razones.append("monto exacto")
        elif diff < 0.05: score += 1; razones.append("monto aprox")
    cuit_c = clean_cuit(comp.get("cuit_origen",""))
    cuit_d = cuit_en_desc(mov_desc)
    if cuit_c and cuit_d and cuit_c[-10:] == cuit_d[-10:]:
        score += 3; razones.append("CUIT")
    c_fecha = normalize_date(comp.get("fecha",""))
    if mov_fecha and c_fecha:
        if mov_fecha == c_fecha: score += 2; razones.append("fecha exacta")
        else:
            try:
                d1 = date.fromisoformat(mov_fecha)
                d2 = date.fromisoformat(c_fecha)
                if abs((d1-d2).days) <= 3: score += 1; razones.append("fecha cercana")
            except: pass
    return score, razones


# ── Filtrar ingresos ──────────────────────────────────────────────

def filtrar_ingresos(df, cols, bank_key):
    cfg  = BANK_CONFIG.get(bank_key, {})
    mask = df[cols["monto"]].apply(normalize_monto) > 0
    if cols["desc"]:
        desc_up = df[cols["desc"]].fillna("").str.strip().str.upper()
        mask_d  = pd.Series([False]*len(df), index=df.index)
        for s in cfg.get("filter_starts",["TRANSFER"]):
            mask_d |= desc_up.str.startswith(s)
        mask &= mask_d
        for cuit_ex in cfg.get("exclude_cuits",[]):
            mask &= ~df[cols["desc"]].fillna("").str.contains(cuit_ex, na=False)
    return df[mask].copy()


# ── Conciliar ─────────────────────────────────────────────────────

def conciliar(planilla_path, comp_paths, bank_key):
    df = pd.read_excel(planilla_path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    HINTS = {
        "fecha": ["fecha","date","dia","día"],
        "monto": ["haber","credito","crédito","importe","monto","ingreso","amount","entrada","credit"],
        "desc":  ["descripcion","descripción","concepto","detalle","movimiento","referencia","desc","glosa"],
    }
    cols = {k: find_col(df.columns.tolist(), v) for k,v in HINTS.items()}
    if not cols["monto"]:
        raise ValueError(f"No se encontro columna de monto. Columnas: {df.columns.tolist()}")
    ing = filtrar_ingresos(df, cols, bank_key)
    comprobantes = []
    for p in comp_paths:
        try:
            data = extract_comprobante(p)
            comprobantes.append({"file": Path(p).name, "data": data})
        except Exception as e:
            comprobantes.append({"file": Path(p).name, "data": {"monto":None,"fecha":None}, "error":str(e)})
    rows = []
    for _, mov in ing.iterrows():
        fecha = normalize_date(mov.get(cols["fecha"],"") if cols["fecha"] else "")
        monto = normalize_monto(mov.get(cols["monto"],0))
        desc  = str(mov.get(cols["desc"],"") if cols["desc"] else "").strip()
        best_score, best_comp, best_razones = 0, None, []
        for c in comprobantes:
            s, r = match_score(fecha, monto, desc, c["data"])
            if s > best_score: best_score, best_comp, best_razones = s, c, r
        verificado = best_score >= UMBRAL_MATCH
        rows.append({
            "fecha": fecha, "descripcion": desc, "importe": round(monto,2),
            "estado": "Verificado" if verificado else "Sin comprobante",
            "comprobante": best_comp["file"] if verificado else "",
            "match": " + ".join(best_razones) if verificado else "",
        })
    comp_usados = {r["comprobante"] for r in rows if r["comprobante"]}
    sin_match = []
    for c in comprobantes:
        if c["file"] not in comp_usados:
            d = c.get("data",{})
            sin_match.append({
                "file":   c["file"],
                "fecha":  normalize_date(d.get("fecha","")),
                "monto":  normalize_monto(d.get("monto") or 0),
                "nombre": d.get("nombre_origen") or "—",
                "banco":  d.get("banco_origen") or "—",
                "cuit":   d.get("cuit_origen") or "—",
            })
    return {"rows": rows, "sin_match": sin_match}


# ── Exportadores ──────────────────────────────────────────────────

def export_excel(rows, sin_match=None):
    wb = Workbook()
    ws = wb.active; ws.title = "Conciliacion"
    verde    = PatternFill("solid", fgColor="C6EFCE")
    amarillo = PatternFill("solid", fgColor="FFEB9C")
    hfill    = PatternFill("solid", fgColor="1F4E79")
    hfont    = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    bfont    = Font(name="Arial", size=10)
    thin     = Side(style="thin", color="BFBFBF")
    border   = Border(left=thin,right=thin,top=thin,bottom=thin)
    ws.append(["Fecha","Descripcion","Importe","Estado","Comprobante","Criterios"])
    for cell in ws[1]:
        cell.fill=hfill; cell.font=hfont
        cell.alignment=Alignment(horizontal="center"); cell.border=border
    for r in rows:
        ws.append([r["fecha"],r["descripcion"],r["importe"],r["estado"],r["comprobante"],r["match"]])
        fill = verde if r["estado"]=="Verificado" else amarillo
        for cell in ws[ws.max_row]:
            cell.fill=fill; cell.font=bfont; cell.border=border
            cell.alignment=Alignment(horizontal="left")
    for col in ws.columns:
        w = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(w+4,55)
    ws.freeze_panes = "A2"
    if sin_match:
        ws2 = wb.create_sheet("Sin match")
        ws2.append(["Archivo","Fecha","Monto","Remitente","Banco","CUIT"])
        for cell in ws2[1]:
            cell.fill=hfill; cell.font=hfont
            cell.alignment=Alignment(horizontal="center"); cell.border=border
        rojo = PatternFill("solid", fgColor="FFCCCC")
        for r in sin_match:
            ws2.append([r["file"],r["fecha"],r["monto"],r["nombre"],r["banco"],r["cuit"]])
            for cell in ws2[ws2.max_row]:
                cell.fill=rojo; cell.font=bfont; cell.border=border
        for col in ws2.columns:
            w = max(len(str(c.value or "")) for c in col)
            ws2.column_dimensions[get_column_letter(col[0].column)].width = min(w+4,50)
        ws2.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

def export_pdf(rows, sin_match=None, bank_label=""):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=2*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    ts = ParagraphStyle("t",parent=styles["Heading1"],fontSize=14,textColor=colors.HexColor("#1F4E79"),spaceAfter=6)
    ss = ParagraphStyle("s",parent=styles["Normal"],fontSize=9,textColor=colors.grey,spaceAfter=14)
    cs = ParagraphStyle("c",parent=styles["Normal"],fontSize=8,leading=11)
    total = len(rows); verif = sum(1 for r in rows if r["estado"]=="Verificado")
    story = [
        Paragraph(f"Conciliacion Bancaria — {bank_label}", ts),
        Paragraph(f"{datetime.now().strftime('%d/%m/%Y %H:%M')}  ·  Total: {total}  ·  Verificados: {verif}  ·  Sin comprobante: {total-verif}", ss),
    ]
    td = [["Fecha","Descripcion","Importe","Estado","Comprobante","Match"]]
    for r in rows:
        td.append([r["fecha"],Paragraph(r["descripcion"][:60],cs),f"${r['importe']:,.0f}",r["estado"],Paragraph(r["comprobante"],cs),r["match"]])
    t = Table(td,colWidths=[2.2*cm,8*cm,2.8*cm,3.2*cm,4.5*cm,5.5*cm],repeatRows=1)
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
    for i,r in enumerate(rows,1):
        fill = colors.HexColor("#C6EFCE") if r["estado"]=="Verificado" else colors.HexColor("#FFEB9C")
        cmds.append(("BACKGROUND",(0,i),(-1,i),fill))
    t.setStyle(TableStyle(cmds)); story.append(t)
    if sin_match:
        story.append(Spacer(1,1*cm))
        story.append(Paragraph("Comprobantes sin coincidencia en el extracto", ts))
        t2d = [["Archivo","Fecha","Monto","Remitente","Banco"]]
        for r in sin_match:
            t2d.append([r["file"],r["fecha"] or "—",f"${r['monto']:,.0f}" if r["monto"] else "—",Paragraph(r["nombre"] or "—",cs),r["banco"] or "—"])
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
    doc.build(story); buf.seek(0)
    return buf


# ── HTML ──────────────────────────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Conciliador Bancario</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>&#x2705;</text></svg>">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#F7F8FA;color:#1A1A2E;font-size:14px}
header{background:#1F4E79;color:#fff;padding:.85rem 2rem;display:flex;align-items:center;gap:12px}
header h1{font-size:17px;font-weight:600}
header span{font-size:13px;opacity:.7}
.container{max-width:1000px;margin:1.5rem auto;padding:0 1.5rem}
.card{background:#fff;border:1px solid #E2E4E8;border-radius:10px;padding:1.5rem;margin-bottom:1.25rem}
.card-title{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#6B7280;margin-bottom:1rem}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem}
.upload-grid{display:grid;grid-template-columns:1fr 1fr;gap:1rem}
select{width:100%;padding:.65rem .85rem;border:1px solid #D1D5DB;border-radius:8px;font-size:14px;outline:none}
select:focus{border-color:#1F4E79}
.drop{border:1.5px dashed #D1D5DB;border-radius:8px;padding:1.5rem;text-align:center;cursor:pointer;transition:.15s;user-select:none}
.drop:hover{background:#F0F4FF;border-color:#93C5FD}
.drop.over{background:#F0F4FF;border-color:#93C5FD}
.file-list{margin-top:.75rem;display:flex;flex-direction:column;gap:6px}
.file-item{display:flex;align-items:center;gap:8px;font-size:12px;padding:5px 10px;background:#F9FAFB;border-radius:6px;border:1px solid #E2E4E8}
.badge{font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px}
.badge-xlsx{background:#EAF3DE;color:#3B6D11}
.badge-pdf{background:#FAECE7;color:#993C1D}
.badge-img{background:#E6F1FB;color:#185FA5}
.fname{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.rm{cursor:pointer;color:#9CA3AF;font-size:15px;line-height:1;padding:0 2px}
.rm:hover{color:#E24B4A}
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
.dot{width:6px;height:6px;border-radius:50%}
.dot-g{background:#1D9E75}.dot-a{background:#EF9F27}
.comp-name{font-size:11px;color:#185FA5;margin-top:2px}
.match-tag{font-size:11px;color:#6B7280}
.sinmatch-title{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#991B1B;margin-bottom:.75rem}
.error-box{background:#FEF2F2;border:1px solid #FCA5A5;border-radius:8px;padding:.85rem 1rem;margin-top:.75rem;font-size:13px;color:#991B1B}
.scroll-wrap{overflow-x:auto}
</style>
</head>
<body>

<header>
  <span style="font-size:22px">&#x2705;</span>
  <div>
    <h1>Conciliador Bancario</h1>
    <span>La Santaniana S.A.</span>
  </div>
</header>

<div class="container">

  <div class="card">
    <div class="card-title">Configuracion</div>
    <div class="form-row">
      <div>
        <div style="font-size:12px;font-weight:600;color:#374151;margin-bottom:4px">Banco</div>
        <select id="bank-sel">
          <option value="Ciudad">Banco Ciudad (Argentina)</option>
          <option value="ITAU">ITAU (Paraguay)</option>
          <option value="UENO">UENO (Paraguay)</option>
        </select>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-title">Archivos</div>
    <div class="upload-grid">

      <div>
        <div class="drop" id="drop-xlsx">
          <div style="font-size:24px;margin-bottom:6px">&#x1F4CA;</div>
          <div style="font-weight:600">Planilla de movimientos</div>
          <div style="font-size:13px;color:#6B7280;margin-top:6px">Click o arrastra el Excel del banco</div>
        </div>
        <input type="file" id="inp-xlsx" accept=".xlsx,.xls,.csv" style="display:none">
        <div class="file-list" id="list-xlsx"></div>
      </div>

      <div>
        <div class="drop" id="drop-comp">
          <div style="font-size:24px;margin-bottom:6px">&#x1F5C2;</div>
          <div style="font-weight:600">Comprobantes de transferencia</div>
          <div style="font-size:13px;color:#6B7280;margin-top:6px">Click o arrastra — PDF, JPG, PNG</div>
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
          <button class="btn-export btn-xlsx" onclick="exportFile('xlsx')">&#x2B07; Excel</button>
          <button class="btn-export btn-pdf"  onclick="exportFile('pdf')">&#x2B07; PDF</button>
        </div>
      </div>
      <div class="scroll-wrap">
        <table><thead><tr>
          <th style="width:90px">Fecha</th>
          <th>Descripcion</th>
          <th style="width:120px;text-align:right">Importe</th>
          <th style="width:145px">Estado</th>
          <th>Comprobante / Criterios</th>
        </tr></thead><tbody id="results-body"></tbody></table>
      </div>
    </div>

    <div id="sinmatch-section" class="hidden">
      <div class="card">
        <div class="results-header">
          <div class="sinmatch-title">&#x26A0; Comprobantes sin coincidencia</div>
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

<script>
var xlsxFile = null;
var compFiles = [];
var lastResults = [];
var lastSinMatch = [];

function ext(n) { return (n || '').split('.').pop().toLowerCase(); }

function badgeClass(n) {
  var e = ext(n);
  if (e === 'xlsx' || e === 'xls' || e === 'csv') return 'badge-xlsx';
  if (e === 'pdf') return 'badge-pdf';
  return 'badge-img';
}

function badgeLabel(n) {
  var e = ext(n);
  if (e === 'xlsx' || e === 'xls') return 'XLSX';
  if (e === 'csv') return 'CSV';
  if (e === 'pdf') return 'PDF';
  return 'IMG';
}

function renderXlsx() {
  var el = document.getElementById('list-xlsx');
  if (!xlsxFile) { el.innerHTML = ''; return; }
  el.innerHTML = '<div class="file-item">' +
    '<span class="badge ' + badgeClass(xlsxFile.name) + '">' + badgeLabel(xlsxFile.name) + '</span>' +
    '<span class="fname">' + xlsxFile.name + '</span>' +
    '<span class="rm" onclick="xlsxFile=null;renderXlsx();check()">x</span>' +
    '</div>';
}

function renderComp() {
  var el = document.getElementById('list-comp');
  el.innerHTML = compFiles.map(function(f, i) {
    return '<div class="file-item">' +
      '<span class="badge ' + badgeClass(f.name) + '">' + badgeLabel(f.name) + '</span>' +
      '<span class="fname">' + f.name + '</span>' +
      '<span class="rm" onclick="compFiles.splice(' + i + ',1);renderComp();check()">x</span>' +
      '</div>';
  }).join('');
}

function check() {
  document.getElementById('run-btn').disabled = !(xlsxFile && compFiles.length > 0);
}

// Click para abrir selector
document.getElementById('drop-xlsx').addEventListener('click', function() {
  document.getElementById('inp-xlsx').click();
});
document.getElementById('drop-comp').addEventListener('click', function() {
  document.getElementById('inp-comp').click();
});

// Cambio en inputs
document.getElementById('inp-xlsx').addEventListener('change', function() {
  if (this.files && this.files.length > 0) {
    xlsxFile = this.files[0];
    renderXlsx();
    check();
  }
});
document.getElementById('inp-comp').addEventListener('change', function() {
  if (this.files && this.files.length > 0) {
    for (var i = 0; i < this.files.length; i++) {
      compFiles.push(this.files[i]);
    }
    renderComp();
    check();
  }
});

// Drag and drop
function setupDrop(dropId, isXlsx) {
  var el = document.getElementById(dropId);
  el.addEventListener('dragover', function(e) {
    e.preventDefault();
    e.stopPropagation();
    el.classList.add('over');
  });
  el.addEventListener('dragleave', function(e) {
    e.stopPropagation();
    el.classList.remove('over');
  });
  el.addEventListener('drop', function(e) {
    e.preventDefault();
    e.stopPropagation();
    el.classList.remove('over');
    var files = Array.from(e.dataTransfer.files);
    if (isXlsx) {
      var f = files.find(function(f) { return ['xlsx','xls','csv'].indexOf(ext(f.name)) >= 0; });
      if (f) { xlsxFile = f; renderXlsx(); check(); }
    } else {
      var valid = files.filter(function(f) { return ['pdf','jpg','jpeg','png'].indexOf(ext(f.name)) >= 0; });
      valid.forEach(function(f) { compFiles.push(f); });
      if (valid.length) { renderComp(); check(); }
    }
  });
}
setupDrop('drop-xlsx', true);
setupDrop('drop-comp', false);

function setProgress(pct, msg) {
  document.getElementById('prog-fill').style.width = pct + '%';
  document.getElementById('status-msg').textContent = msg;
}

function showError(msg) {
  document.getElementById('error-area').innerHTML = '<div class="error-box">' + msg + '</div>';
  document.getElementById('prog-wrap').classList.add('hidden');
  document.getElementById('run-btn').disabled = false;
}

async function runConciliation() {
  document.getElementById('error-area').innerHTML = '';
  document.getElementById('results-section').classList.add('hidden');
  document.getElementById('prog-wrap').classList.remove('hidden');
  document.getElementById('run-btn').disabled = true;
  setProgress(5, 'Subiendo archivos...');

  var fd = new FormData();
  fd.append('planilla', xlsxFile);
  for (var i = 0; i < compFiles.length; i++) {
    fd.append('comprobantes', compFiles[i]);
  }
  fd.append('bank', document.getElementById('bank-sel').value);

  try {
    setProgress(15, 'Procesando con IA — puede tardar 1-2 minutos...');
    var resp = await fetch('/conciliar', { method: 'POST', body: fd });
    var data = await resp.json();
    if (data.error) { showError(data.error); return; }
    setProgress(100, 'Conciliacion completada.');
    lastResults = data.rows;
    lastSinMatch = data.sin_match || [];
    renderResults(data.rows);
    renderSinMatch(data.sin_match || []);
  } catch(e) {
    showError('Error: ' + e.message);
  } finally {
    document.getElementById('run-btn').disabled = false;
  }
}

function renderResults(rows) {
  var total = rows.length;
  var verif = rows.filter(function(r) { return r.estado === 'Verificado'; }).length;
  var totalM = rows.reduce(function(a, r) { return a + r.importe; }, 0);

  document.getElementById('stats-grid').innerHTML =
    '<div class="stat"><div class="stat-label">Transferencias</div><div class="stat-value">' + total + '</div></div>' +
    '<div class="stat"><div class="stat-label">Verificados</div><div class="stat-value green">' + verif + '</div></div>' +
    '<div class="stat"><div class="stat-label">Sin comprobante</div><div class="stat-value amber">' + (total-verif) + '</div></div>' +
    '<div class="stat"><div class="stat-label">Total importe</div><div class="stat-value" style="font-size:15px">$' + totalM.toLocaleString('es-AR',{minimumFractionDigits:0}) + '</div></div>';

  document.getElementById('results-body').innerHTML = rows.map(function(r) {
    var chip = r.estado === 'Verificado'
      ? '<div class="chip chip-v"><div class="dot dot-g"></div>Verificado</div>'
      : '<div class="chip chip-m"><div class="dot dot-a"></div>Sin comprobante</div>';
    var comp = r.comprobante ? '<div class="comp-name">' + r.comprobante + '</div>' : '';
    var match = r.match ? '<div class="match-tag">' + r.match + '</div>' : '&#x2014;';
    return '<tr class="' + (r.estado === 'Verificado' ? 'verified' : 'missing') + '">' +
      '<td style="color:#6B7280">' + r.fecha + '</td>' +
      '<td>' + r.descripcion + '</td>' +
      '<td style="text-align:right;font-weight:600">$' + r.importe.toLocaleString('es-AR',{minimumFractionDigits:2}) + '</td>' +
      '<td>' + chip + '</td>' +
      '<td>' + comp + match + '</td>' +
      '</tr>';
  }).join('');

  document.getElementById('results-section').classList.remove('hidden');
}

function renderSinMatch(items) {
  var s = document.getElementById('sinmatch-section');
  if (!items || items.length === 0) { s.classList.add('hidden'); return; }
  document.getElementById('sinmatch-count').textContent = items.length + ' comprobante(s)';
  document.getElementById('sinmatch-body').innerHTML = items.map(function(r) {
    var monto = r.monto ? '$' + r.monto.toLocaleString('es-AR',{minimumFractionDigits:0}) : '&#x2014;';
    return '<tr class="sinmatch">' +
      '<td style="font-size:12px">' + r.file + '</td>' +
      '<td style="color:#6B7280">' + r.fecha + '</td>' +
      '<td style="text-align:right;font-weight:600">' + monto + '</td>' +
      '<td>' + r.nombre + '</td><td>' + r.banco + '</td>' +
      '<td style="font-family:monospace;font-size:11px">' + r.cuit + '</td>' +
      '</tr>';
  }).join('');
  s.classList.remove('hidden');
}

function exportFile(fmt) {
  if (!lastResults.length) return;
  var fd = new FormData();
  fd.append('rows', JSON.stringify(lastResults));
  fd.append('sin_match', JSON.stringify(lastSinMatch));
  fd.append('format', fmt);
  fd.append('bank_label', document.getElementById('bank-sel').options[document.getElementById('bank-sel').selectedIndex].text);
  fetch('/exportar', { method: 'POST', body: fd })
    .then(function(r) { return r.blob(); })
    .then(function(blob) {
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = fmt === 'xlsx' ? 'conciliacion.xlsx' : 'conciliacion.pdf';
      a.click();
      URL.revokeObjectURL(url);
    });
}
</script>
</body>
</html>"""


# ── Rutas ─────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/conciliar", methods=["POST"])
def conciliar_route():
    if API_KEY.startswith("sk-ant-..."):
        return jsonify({"error": "API key no configurada. Agregala en Railway → Variables → ANTHROPIC_API_KEY"})
    try:
        planilla   = request.files["planilla"]
        comps      = request.files.getlist("comprobantes")
        bank_key   = request.form.get("bank","Ciudad")
        plan_path  = os.path.join(UPLOAD_FOLDER, "planilla" + Path(planilla.filename).suffix)
        planilla.save(plan_path)
        comp_paths = []
        for f in comps:
            p = os.path.join(UPLOAD_FOLDER, f.filename)
            f.save(p); comp_paths.append(p)
        result = conciliar(plan_path, comp_paths, bank_key)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/exportar", methods=["POST"])
def exportar_route():
    rows      = json.loads(request.form["rows"])
    sin_match = json.loads(request.form.get("sin_match","[]"))
    fmt        = request.form["format"]
    bank_label = request.form.get("bank_label","")
    if fmt == "xlsx":
        buf = export_excel(rows, sin_match)
        return send_file(buf, download_name="conciliacion.xlsx", as_attachment=True,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        buf = export_pdf(rows, sin_match, bank_label)
        return send_file(buf, download_name="conciliacion.pdf", as_attachment=True,
                         mimetype="application/pdf")


# ── Main ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    host = "0.0.0.0" if os.environ.get("RAILWAY_ENVIRONMENT") else "127.0.0.1"
    if host == "127.0.0.1":
        print("=" * 50)
        print("  Conciliador Bancario")
        print("  http://localhost:5000")
        print("=" * 50)
    app.run(debug=False, host=host, port=port)

