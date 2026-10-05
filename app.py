"""
Conciliador Bancario Web
========================
Correr local: python app.py  →  http://localhost:5000
Railway:      gunicorn app:app --timeout 120
"""

import os, re, json, base64, io, sqlite3, hashlib
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

# CONFIG
API_KEY       = os.environ.get("ANTHROPIC_API_KEY", "sk-ant-...")
UPLOAD_FOLDER = "uploads_tmp"
DB_PATH       = os.environ.get("DB_PATH", "conciliador.db")
UMBRAL_MATCH  = 4

BANK_CONFIG = {
    "Ciudad": {"label":"Banco Ciudad (Argentina)","filter_starts":["TRANSFER"],"exclude_cuits":["30708635754","30711747709"]},
    "ITAU":   {"label":"ITAU (Paraguay)",          "filter_starts":["TRANSFER"],"exclude_cuits":[]},
    "UENO":   {"label":"UENO (Paraguay)",           "filter_starts":["TRANSFER"],"exclude_cuits":[]},
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# DB
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS conciliaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            banco TEXT, total_comp INTEGER, verificados INTEGER, sin_match INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS comp_usados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_hash TEXT UNIQUE, nombre_archivo TEXT,
            monto REAL, fecha_comp TEXT, nombre_origen TEXT,
            conciliacion_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    conn.commit(); conn.close()

def hash_file(path):
    h = hashlib.md5()
    with open(path,"rb") as f: h.update(f.read())
    return h.hexdigest()

def check_duplicados(comp_paths):
    conn = get_db(); dups = []
    for p in comp_paths:
        fh = hash_file(p)
        row = conn.execute("SELECT nombre_archivo,conciliacion_id FROM comp_usados WHERE file_hash=?",(fh,)).fetchone()
        if row: dups.append({"archivo":Path(p).name,"hash":fh,"conciliacion_id":row["conciliacion_id"]})
    conn.close(); return dups

def guardar_conciliacion(banco, comp_data_list, verificados):
    conn = get_db()
    cur = conn.execute("INSERT INTO conciliaciones (banco,total_comp,verificados,sin_match) VALUES (?,?,?,?)",
        (banco, len(comp_data_list), verificados, len(comp_data_list)-verificados))
    cid = cur.lastrowid
    for item in comp_data_list:
        try:
            conn.execute("INSERT OR IGNORE INTO comp_usados (file_hash,nombre_archivo,monto,fecha_comp,nombre_origen,conciliacion_id) VALUES (?,?,?,?,?,?)",
                (item["hash"],item["file"],item.get("monto"),item.get("fecha"),item.get("nombre"),cid))
        except: pass
    conn.commit(); conn.close(); return cid

# Helpers
def find_col(keys, hints):
    lower = [k.lower().strip() for k in keys]
    for h in hints:
        for i,k in enumerate(lower):
            if h in k: return keys[i]
    return None

def normalize_monto(v):
    if v is None or v=="": return 0.0
    if isinstance(v,(int,float)): return float(v)
    s = str(v).strip(); neg = s.startswith("-")
    s = re.sub(r"[^\d,\.]","",s)
    if not s: return 0.0
    if "," in s and "." in s: r=float(s.replace(".","").replace(",","."))
    elif "," in s: r=float(s.replace(",","."))
    else: r=float(s) if s else 0.0
    return -r if neg else r

def normalize_date(v):
    if not v: return ""
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$",s)
    if m:
        y=("20"+m.group(3)) if len(m.group(3))==2 else m.group(3)
        return f"{y}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
    m = re.match(r"^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$",s)
    if m: return f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
    if "T" in s: return s.split("T")[0]
    return s[:10]

def clean_cuit(c): return re.sub(r"[^\d]","",str(c)) if c else ""
def cuit_en_desc(desc):
    m = re.search(r"(?:^|\s)(\d{10,11})[\s\-]",str(desc))
    return m.group(1) if m else ""
def file_to_b64(path):
    with open(path,"rb") as f: return base64.standard_b64encode(f.read()).decode()
def get_media_type(path):
    return {".pdf":"application/pdf",".jpg":"image/jpeg",".jpeg":"image/jpeg",".png":"image/png"}.get(Path(path).suffix.lower(),"")

# IA
def extract_comprobante(path):
    client = anthropic.Anthropic(api_key=API_KEY)
    b64=file_to_b64(path); mt=get_media_type(path)
    is_img = Path(path).suffix.lower() in {".jpg",".jpeg",".png"}
    resp = client.messages.create(model="claude-opus-4-5",max_tokens=700,
        messages=[{"role":"user","content":[
            {"type":"image" if is_img else "document","source":{"type":"base64","media_type":mt,"data":b64}},
            {"type":"text","text":"Comprobante de transferencia bancaria. Devolvé SOLO JSON con: fecha (DD/MM/YYYY), monto (numero sin simbolo ni puntos de miles), moneda, cuit_origen (solo digitos), nombre_origen, banco_origen, referencia. Null si no aparece. Sin backticks."}
        ]}])
    text="".join(b.text for b in resp.content if hasattr(b,"text"))
    try: return json.loads(re.sub(r"```json|```","",text).strip())
    except: return {"monto":None,"fecha":None,"error":"parse_error"}

# Match
def match_score(mov_fecha,mov_monto,mov_desc,comp):
    score,razones=0,[]
    c_monto=normalize_monto(comp.get("monto"))
    if c_monto and mov_monto:
        diff=abs(mov_monto-c_monto)/max(mov_monto,c_monto)
        if diff<0.01: score+=3; razones.append("monto exacto")
        elif diff<0.05: score+=1; razones.append("monto aprox")
    cuit_c=clean_cuit(comp.get("cuit_origen",""))
    cuit_d=cuit_en_desc(mov_desc)
    if cuit_c and cuit_d and cuit_c[-10:]==cuit_d[-10:]: score+=3; razones.append("CUIT")
    c_fecha=normalize_date(comp.get("fecha",""))
    if mov_fecha and c_fecha:
        if mov_fecha==c_fecha: score+=2; razones.append("fecha exacta")
        else:
            try:
                d1=date.fromisoformat(mov_fecha); d2=date.fromisoformat(c_fecha)
                if abs((d1-d2).days)<=3: score+=1; razones.append("fecha cercana")
            except: pass
    return score,razones

def filtrar_ingresos(df,cols,bank_key):
    cfg=BANK_CONFIG.get(bank_key,{})
    mask=df[cols["monto"]].apply(normalize_monto)>0
    if cols["desc"]:
        desc_up=df[cols["desc"]].fillna("").str.strip().str.upper()
        mask_d=pd.Series([False]*len(df),index=df.index)
        for s in cfg.get("filter_starts",["TRANSFER"]): mask_d|=desc_up.str.startswith(s)
        mask&=mask_d
        for cx in cfg.get("exclude_cuits",[]): mask&=~df[cols["desc"]].fillna("").str.contains(cx,na=False)
    return df[mask].copy()

def conciliar(planilla_path,comp_paths,bank_key):
    df=pd.read_excel(planilla_path,dtype=str)
    df.columns=[str(c).strip() for c in df.columns]
    HINTS={"fecha":["fecha","date","dia","día"],"monto":["haber","credito","crédito","importe","monto","ingreso","amount","entrada","credit"],"desc":["descripcion","descripción","concepto","detalle","movimiento","referencia","desc","glosa"]}
    cols={k:find_col(df.columns.tolist(),v) for k,v in HINTS.items()}
    if not cols["monto"]: raise ValueError(f"No se encontro columna de monto. Columnas: {df.columns.tolist()}")
    ing=filtrar_ingresos(df,cols,bank_key)
    comprobantes=[]
    for p in comp_paths:
        try:
            data=extract_comprobante(p)
            comprobantes.append({"file":Path(p).name,"hash":hash_file(p),"data":data,
                "monto":normalize_monto(data.get("monto")),"fecha":normalize_date(data.get("fecha","")),
                "nombre":data.get("nombre_origen",""),"banco_origen":data.get("banco_origen",""),"cuit":data.get("cuit_origen","")})
        except Exception as e:
            comprobantes.append({"file":Path(p).name,"hash":hash_file(p),"data":{"monto":None,"fecha":None},"error":str(e)})
    rows=[]
    for _,mov in ing.iterrows():
        fecha=normalize_date(mov.get(cols["fecha"],"") if cols["fecha"] else "")
        monto=normalize_monto(mov.get(cols["monto"],0))
        desc=str(mov.get(cols["desc"],"") if cols["desc"] else "").strip()
        best_score,best_comp,best_razones=0,None,[]
        for c in comprobantes:
            s,r=match_score(fecha,monto,desc,c["data"])
            if s>best_score: best_score,best_comp,best_razones=s,c,r
        verificado=best_score>=UMBRAL_MATCH
        rows.append({"fecha":fecha,"descripcion":desc,"importe":round(monto,2),
            "estado":"Verificado" if verificado else "Sin comprobante",
            "comprobante":best_comp["file"] if verificado else "","match":" + ".join(best_razones) if verificado else ""})
    comp_verificados={r["comprobante"] for r in rows if r["comprobante"]}
    comp_result=[]
    for c in comprobantes:
        verificado=c["file"] in comp_verificados
        mov_match=next((r for r in rows if r["comprobante"]==c["file"]),None)
        comp_result.append({"file":c["file"],"hash":c["hash"],"fecha":c.get("fecha",""),"monto":c.get("monto",0),
            "nombre":c.get("nombre","—"),"banco_origen":c.get("banco_origen","—"),"cuit":c.get("cuit","—"),
            "verificado":verificado,"match":mov_match["match"] if mov_match else ""})
    sin_match=[c for c in comp_result if not c["verificado"]]
    return {"rows":rows,"comp_result":comp_result,"sin_match":sin_match,"total_comp":len(comprobantes),"verificados":sum(1 for c in comp_result if c["verificado"])}

# Exportadores
def export_excel(rows,comp_result,sin_match,cid,bank_label):
    wb=Workbook(); ws=wb.active; ws.title="Comprobantes"
    verde=PatternFill("solid",fgColor="C6EFCE"); amarillo=PatternFill("solid",fgColor="FFEB9C")
    rojo=PatternFill("solid",fgColor="FFCCCC"); hfill=PatternFill("solid",fgColor="1F4E79")
    hfont=Font(bold=True,color="FFFFFF",name="Arial",size=10); bfont=Font(name="Arial",size=10)
    thin=Side(style="thin",color="BFBFBF"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
    ws.append([f"Conciliacion N {cid} - {bank_label} - {datetime.now().strftime('%d/%m/%Y')}"])
    ws.merge_cells("A1:H1"); ws["A1"].font=Font(bold=True,name="Arial",size=11,color="1F4E79")
    ws.append([])
    ws.append(["Archivo","Fecha","Monto","Remitente","Banco Origen","CUIT","Estado","Criterios"])
    for cell in ws[3]:
        cell.fill=hfill; cell.font=hfont; cell.alignment=Alignment(horizontal="center"); cell.border=border
    for c in comp_result:
        ws.append([c["file"],c["fecha"],c["monto"] or 0,c["nombre"],c["banco_origen"],c["cuit"],
            "Verificado" if c["verificado"] else "Sin coincidencia",c["match"]])
        fill=verde if c["verificado"] else rojo
        for cell in ws[ws.max_row]: cell.fill=fill; cell.font=bfont; cell.border=border; cell.alignment=Alignment(horizontal="left")
    for col in ws.columns:
        w=max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width=min(w+4,50)
    ws.freeze_panes="A4"
    ws2=wb.create_sheet("Extracto bancario")
    ws2.append(["Fecha","Descripcion","Importe","Estado","Comprobante","Criterios"])
    for cell in ws2[1]: cell.fill=hfill; cell.font=hfont; cell.alignment=Alignment(horizontal="center"); cell.border=border
    for r in rows:
        ws2.append([r["fecha"],r["descripcion"],r["importe"],r["estado"],r["comprobante"],r["match"]])
        fill=verde if r["estado"]=="Verificado" else amarillo
        for cell in ws2[ws2.max_row]: cell.fill=fill; cell.font=bfont; cell.border=border
    for col in ws2.columns:
        w=max(len(str(c.value or "")) for c in col)
        ws2.column_dimensions[get_column_letter(col[0].column)].width=min(w+4,55)
    ws2.freeze_panes="A2"
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def export_pdf(comp_result,cid,bank_label):
    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=landscape(A4),leftMargin=1.5*cm,rightMargin=1.5*cm,topMargin=2*cm,bottomMargin=1.5*cm)
    styles=getSampleStyleSheet()
    ts=ParagraphStyle("t",parent=styles["Heading1"],fontSize=14,textColor=colors.HexColor("#1F4E79"),spaceAfter=4)
    ss=ParagraphStyle("s",parent=styles["Normal"],fontSize=9,textColor=colors.grey,spaceAfter=16)
    cs=ParagraphStyle("c",parent=styles["Normal"],fontSize=8,leading=11)
    total=len(comp_result); verif=sum(1 for c in comp_result if c["verificado"])
    story=[
        Paragraph(f"Conciliacion N{chr(176)} {cid} - {bank_label}",ts),
        Paragraph(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Ingresados: {total}  |  Verificados: {verif}  |  Sin coincidencia: {total-verif}",ss),
    ]
    verif_list=[c for c in comp_result if c["verificado"]]
    if verif_list:
        story.append(Paragraph("Comprobantes verificados",ParagraphStyle("h2",parent=styles["Heading2"],fontSize=11,textColor=colors.HexColor("#166534"),spaceAfter=6)))
        td=[["Archivo","Fecha","Monto","Remitente","Criterios de match"]]
        for c in verif_list:
            td.append([Paragraph(c["file"],cs),c["fecha"] or "-",f"${c['monto']:,.0f}" if c.get("monto") else "-",Paragraph(c["nombre"] or "-",cs),c["match"] or "-"])
        t=Table(td,colWidths=[5*cm,2.5*cm,3*cm,7*cm,8*cm],repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#166534")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("FONTNAME",(0,1),(-1,-1),"Helvetica"),
            ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#C6EFCE")),("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),5),
        ])); story.append(t)
    sin_list=[c for c in comp_result if not c["verificado"]]
    if sin_list:
        story.append(Spacer(1,0.7*cm))
        story.append(Paragraph("Comprobantes sin coincidencia",ParagraphStyle("h2b",parent=styles["Heading2"],fontSize=11,textColor=colors.HexColor("#991B1B"),spaceAfter=6)))
        td2=[["Archivo","Fecha","Monto","Remitente","Banco origen"]]
        for c in sin_list:
            td2.append([Paragraph(c["file"],cs),c["fecha"] or "-",f"${c['monto']:,.0f}" if c.get("monto") else "-",Paragraph(c["nombre"] or "-",cs),c["banco_origen"] or "-"])
        t2=Table(td2,colWidths=[5*cm,2.5*cm,3*cm,7*cm,8*cm],repeatRows=1)
        t2.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#C0392B")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("FONTNAME",(0,1),(-1,-1),"Helvetica"),
            ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#FFCCCC")),("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),5),
        ])); story.append(t2)
    doc.build(story); buf.seek(0); return buf

HTML_PAGE = '<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width,initial-scale=1">\n<title>Conciliador Bancario</title>\n<link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>&#x2705;</text></svg>">\n<style>\n*{box-sizing:border-box;margin:0;padding:0}\nbody{font-family:Arial,sans-serif;background:#F7F8FA;color:#1A1A2E;font-size:14px}\nheader{background:#1F4E79;color:#fff;padding:.85rem 2rem;display:flex;align-items:center;gap:12px}\nheader h1{font-size:17px;font-weight:600}\nheader span{font-size:13px;opacity:.7}\n.container{max-width:1000px;margin:1.5rem auto;padding:0 1.5rem}\n.card{background:#fff;border:1px solid #E2E4E8;border-radius:10px;padding:1.5rem;margin-bottom:1.25rem}\n.card-title{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#6B7280;margin-bottom:1rem}\n.form-row{display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem}\n.upload-grid{display:grid;grid-template-columns:1fr 1fr;gap:1rem}\nselect{width:100%;padding:.65rem .85rem;border:1px solid #D1D5DB;border-radius:8px;font-size:14px;outline:none}\nselect:focus{border-color:#1F4E79}\n.drop{border:1.5px dashed #D1D5DB;border-radius:8px;padding:1.5rem;text-align:center;cursor:pointer;transition:.15s;user-select:none}\n.drop:hover,.drop.over{background:#F0F4FF;border-color:#93C5FD}\n.file-list{margin-top:.75rem;display:flex;flex-direction:column;gap:6px}\n.file-item{display:flex;align-items:center;gap:8px;font-size:12px;padding:5px 10px;background:#F9FAFB;border-radius:6px;border:1px solid #E2E4E8}\n.badge{font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px}\n.badge-xlsx{background:#EAF3DE;color:#3B6D11}\n.badge-pdf{background:#FAECE7;color:#993C1D}\n.badge-img{background:#E6F1FB;color:#185FA5}\n.fname{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}\n.rm{cursor:pointer;color:#9CA3AF;font-size:15px;line-height:1;padding:0 2px}\n.rm:hover{color:#E24B4A}\n.run-btn{width:100%;padding:.85rem;font-size:15px;font-weight:600;background:#1F4E79;color:#fff;border:none;border-radius:8px;cursor:pointer;transition:.15s}\n.run-btn:hover:not(:disabled){background:#163A5F}\n.run-btn:disabled{opacity:.45;cursor:not-allowed}\n.progress-bar{height:5px;background:#E2E4E8;border-radius:3px;overflow:hidden;margin-top:.75rem}\n.progress-fill{height:100%;background:#1D9E75;border-radius:3px;transition:width .4s}\n.status-msg{font-size:13px;color:#6B7280;text-align:center;margin-top:.5rem;min-height:18px}\n.hidden{display:none!important}\n.stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:1.25rem}\n.stat{background:#F9FAFB;border:1px solid #E2E4E8;border-radius:8px;padding:.85rem 1rem}\n.stat-label{font-size:11px;color:#6B7280;margin-bottom:4px}\n.stat-value{font-size:22px;font-weight:700}\n.stat-num{font-size:15px;color:#1F4E79;font-weight:700}\n.green{color:#1D9E75}.amber{color:#854F0B}\n.results-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:.85rem}\n.export-row{display:flex;gap:8px}\n.btn-export{padding:.45rem 1rem;font-size:13px;font-weight:500;border:1px solid #E2E4E8;background:#fff;border-radius:6px;cursor:pointer}\n.btn-export:hover{background:#F0F4FF}\n.btn-xlsx{border-color:#3B6D11;color:#3B6D11}\n.btn-pdf{border-color:#993C1D;color:#993C1D}\ntable{width:100%;border-collapse:collapse;font-size:13px}\nth{text-align:left;padding:8px 10px;font-size:11px;font-weight:600;color:#6B7280;border-bottom:2px solid #E2E4E8;white-space:nowrap}\ntd{padding:8px 10px;border-bottom:1px solid #E2E4E8;vertical-align:middle}\ntr:last-child td{border-bottom:none}\ntr.verified td{background:#C6EFCE}\ntr.sinmatch td{background:#FFCCCC}\n.chip{display:inline-flex;align-items:center;gap:4px;font-size:11px;font-weight:600;padding:3px 9px;border-radius:4px}\n.chip-v{background:rgba(0,0,0,.08);color:#166534}\n.chip-r{background:rgba(0,0,0,.08);color:#991B1B}\n.dot{width:6px;height:6px;border-radius:50%}\n.dot-g{background:#1D9E75}.dot-r{background:#E24B4A}\n.match-tag{font-size:11px;color:#6B7280}\n.sinmatch-title{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#991B1B;margin-bottom:.75rem}\n.error-box{background:#FEF2F2;border:1px solid #FCA5A5;border-radius:8px;padding:.85rem 1rem;margin-top:.75rem;font-size:13px;color:#991B1B}\n.scroll-wrap{overflow-x:auto}\n.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;z-index:1000}\n.modal{background:#fff;border-radius:12px;padding:2rem;max-width:520px;width:90%;box-shadow:0 8px 32px rgba(0,0,0,.2)}\n.modal h2{font-size:16px;color:#991B1B;margin-bottom:.5rem}\n.modal p{font-size:13px;color:#6B7280;margin-bottom:1rem}\n.modal-list{background:#FEF2F2;border-radius:8px;padding:.75rem 1rem;margin-bottom:1.25rem;font-size:13px}\n.modal-list li{padding:3px 0;color:#374151}\n.modal-btns{display:flex;gap:.75rem;justify-content:flex-end}\n.btn-cancel{padding:.6rem 1.25rem;border:1px solid #E2E4E8;border-radius:7px;background:#fff;cursor:pointer;font-size:13px}\n.btn-force{padding:.6rem 1.25rem;border:none;border-radius:7px;background:#991B1B;color:#fff;cursor:pointer;font-size:13px;font-weight:600}\n</style>\n</head>\n<body>\n<div class="modal-overlay hidden" id="modal-dup">\n  <div class="modal">\n    <h2>&#x26A0; Comprobantes ya utilizados</h2>\n    <p>Los siguientes comprobantes ya fueron procesados en una conciliacion anterior:</p>\n    <ul class="modal-list" id="modal-list"></ul>\n    <p style="font-size:13px;color:#374151">Desea incluirlos de todas formas en esta conciliacion?</p>\n    <div class="modal-btns">\n      <button class="btn-cancel" onclick="cerrarModal()">Cancelar</button>\n      <button class="btn-force" onclick="confirmarDuplicados()">Si, incluir de todas formas</button>\n    </div>\n  </div>\n</div>\n<header>\n  <span style="font-size:22px">&#x2705;</span>\n  <div><h1>Conciliador Bancario</h1><span>La Santaniana S.A.</span></div>\n</header>\n<div class="container">\n  <div class="card">\n    <div class="card-title">Configuracion</div>\n    <div class="form-row">\n      <div>\n        <div style="font-size:12px;font-weight:600;color:#374151;margin-bottom:4px">Banco</div>\n        <select id="bank-sel">\n          <option value="Ciudad">Banco Ciudad (Argentina)</option>\n          <option value="ITAU">ITAU (Paraguay)</option>\n          <option value="UENO">UENO (Paraguay)</option>\n        </select>\n      </div>\n    </div>\n  </div>\n  <div class="card">\n    <div class="card-title">Archivos</div>\n    <div class="upload-grid">\n      <div>\n        <div class="drop" id="drop-xlsx">\n          <div style="font-size:24px;margin-bottom:6px">&#x1F4CA;</div>\n          <div style="font-weight:600">Planilla de movimientos</div>\n          <div style="font-size:13px;color:#6B7280;margin-top:6px">Click o arrastra el Excel del banco</div>\n        </div>\n        <input type="file" id="inp-xlsx" accept=".xlsx,.xls,.csv" style="display:none">\n        <div class="file-list" id="list-xlsx"></div>\n      </div>\n      <div>\n        <div class="drop" id="drop-comp">\n          <div style="font-size:24px;margin-bottom:6px">&#x1F5C2;</div>\n          <div style="font-weight:600">Comprobantes de transferencia</div>\n          <div style="font-size:13px;color:#6B7280;margin-top:6px">Click o arrastra — PDF, JPG, PNG</div>\n        </div>\n        <input type="file" id="inp-comp" accept=".pdf,.jpg,.jpeg,.png" multiple style="display:none">\n        <div class="file-list" id="list-comp"></div>\n      </div>\n    </div>\n  </div>\n  <button class="run-btn" id="run-btn" disabled onclick="runConciliation(false)">Conciliar transferencias</button>\n  <div class="progress-bar hidden" id="prog-wrap"><div class="progress-fill" id="prog-fill" style="width:0%"></div></div>\n  <div class="status-msg" id="status-msg"></div>\n  <div id="error-area"></div>\n  <div id="results-section" class="hidden" style="margin-top:1.5rem">\n    <div class="stats-grid" id="stats-grid"></div>\n    <div class="card">\n      <div class="results-header">\n        <div class="card-title" style="margin-bottom:0">Comprobantes</div>\n        <div class="export-row">\n          <button class="btn-export btn-xlsx" onclick="exportFile(\'xlsx\')">&#x2B07; Excel</button>\n          <button class="btn-export btn-pdf"  onclick="exportFile(\'pdf\')">&#x2B07; PDF</button>\n        </div>\n      </div>\n      <div class="scroll-wrap">\n        <table><thead><tr>\n          <th>Archivo</th><th style="width:90px">Fecha</th>\n          <th style="width:110px;text-align:right">Monto</th>\n          <th>Remitente</th><th style="width:145px">Estado</th><th>Criterios</th>\n        </tr></thead><tbody id="comp-body"></tbody></table>\n      </div>\n    </div>\n    <div id="sinmatch-section" class="hidden">\n      <div class="card">\n        <div class="sinmatch-title">&#x26A0; Comprobantes sin coincidencia en el extracto</div>\n        <div class="scroll-wrap">\n          <table><thead><tr>\n            <th>Archivo</th><th style="width:90px">Fecha</th>\n            <th style="width:110px;text-align:right">Monto</th>\n            <th>Remitente</th><th>Banco</th><th>CUIT</th>\n          </tr></thead><tbody id="sinmatch-body"></tbody></table>\n        </div>\n      </div>\n    </div>\n  </div>\n</div>\n<script>\nvar xlsxFile=null,compFiles=[],lastCompResult=[],lastRows=[],lastSinMatch=[],lastCid=null,lastBankLabel="";\nfunction ext(n){return(n||"").split(".").pop().toLowerCase();}\nfunction badgeClass(n){var e=ext(n);if(e==="xlsx"||e==="xls"||e==="csv")return"badge-xlsx";if(e==="pdf")return"badge-pdf";return"badge-img";}\nfunction badgeLabel(n){var e=ext(n);if(e==="xlsx"||e==="xls")return"XLSX";if(e==="csv")return"CSV";if(e==="pdf")return"PDF";return"IMG";}\nfunction renderXlsx(){var el=document.getElementById("list-xlsx");if(!xlsxFile){el.innerHTML="";return;}el.innerHTML="<div class=\\"file-item\\"><span class=\\"badge "+badgeClass(xlsxFile.name)+"\\">"+badgeLabel(xlsxFile.name)+"</span><span class=\\"fname\\">"+xlsxFile.name+"</span><span class=\\"rm\\" onclick=\\"xlsxFile=null;renderXlsx();check()\\">x</span></div>";}\nfunction renderComp(){var el=document.getElementById("list-comp");el.innerHTML=compFiles.map(function(f,i){return"<div class=\\"file-item\\"><span class=\\"badge "+badgeClass(f.name)+"\\">"+badgeLabel(f.name)+"</span><span class=\\"fname\\">"+f.name+"</span><span class=\\"rm\\" onclick=\\"compFiles.splice("+i+",1);renderComp();check()\\">x</span></div>";}).join("");}\nfunction check(){document.getElementById("run-btn").disabled=!(xlsxFile&&compFiles.length>0);}\ndocument.getElementById("drop-xlsx").addEventListener("click",function(){document.getElementById("inp-xlsx").click();});\ndocument.getElementById("drop-comp").addEventListener("click",function(){document.getElementById("inp-comp").click();});\ndocument.getElementById("inp-xlsx").addEventListener("change",function(){if(this.files&&this.files.length>0){xlsxFile=this.files[0];renderXlsx();check();}});\ndocument.getElementById("inp-comp").addEventListener("change",function(){if(this.files&&this.files.length>0){for(var i=0;i<this.files.length;i++)compFiles.push(this.files[i]);renderComp();check();}});\nfunction setupDrop(id,isXlsx){var el=document.getElementById(id);el.addEventListener("dragover",function(e){e.preventDefault();e.stopPropagation();el.classList.add("over");});el.addEventListener("dragleave",function(e){e.stopPropagation();el.classList.remove("over");});el.addEventListener("drop",function(e){e.preventDefault();e.stopPropagation();el.classList.remove("over");var files=Array.from(e.dataTransfer.files);if(isXlsx){var f=files.find(function(f){return["xlsx","xls","csv"].indexOf(ext(f.name))>=0;});if(f){xlsxFile=f;renderXlsx();check();}}else{var valid=files.filter(function(f){return["pdf","jpg","jpeg","png"].indexOf(ext(f.name))>=0;});valid.forEach(function(f){compFiles.push(f);});if(valid.length){renderComp();check();}}});}\nsetupDrop("drop-xlsx",true);setupDrop("drop-comp",false);\nfunction setProgress(pct,msg){document.getElementById("prog-fill").style.width=pct+"%";document.getElementById("status-msg").textContent=msg;}\nfunction showError(msg){document.getElementById("error-area").innerHTML="<div class=\\"error-box\\">"+msg+"</div>";document.getElementById("prog-wrap").classList.add("hidden");document.getElementById("run-btn").disabled=false;}\nasync function runConciliation(force){\n  document.getElementById("error-area").innerHTML="";\n  document.getElementById("results-section").classList.add("hidden");\n  document.getElementById("prog-wrap").classList.remove("hidden");\n  document.getElementById("run-btn").disabled=true;\n  setProgress(5,"Subiendo archivos...");\n  var bankSel=document.getElementById("bank-sel");\n  lastBankLabel=bankSel.options[bankSel.selectedIndex].text;\n  var fd=new FormData();\n  fd.append("planilla",xlsxFile);\n  for(var i=0;i<compFiles.length;i++)fd.append("comprobantes",compFiles[i]);\n  fd.append("bank",bankSel.value);\n  fd.append("force",force?"true":"false");\n  try{\n    setProgress(15,"Procesando con IA — puede tardar 1-2 minutos...");\n    var resp=await fetch("/conciliar",{method:"POST",body:fd});\n    var data=await resp.json();\n    if(data.duplicados){document.getElementById("prog-wrap").classList.add("hidden");document.getElementById("status-msg").textContent="";document.getElementById("run-btn").disabled=false;mostrarModal(data.duplicados);return;}\n    if(data.error){showError(data.error);return;}\n    setProgress(100,"Conciliacion N° "+data.conciliacion_id+" completada.");\n    lastRows=data.rows;lastCompResult=data.comp_result;lastSinMatch=data.sin_match||[];lastCid=data.conciliacion_id;\n    renderStats(data);renderCompResult(data.comp_result);renderSinMatch(data.sin_match||[]);\n  }catch(e){showError("Error: "+e.message);}\n  finally{document.getElementById("run-btn").disabled=false;}\n}\nfunction mostrarModal(dups){var lista=document.getElementById("modal-list");lista.innerHTML=dups.map(function(d){return"<li>"+d.archivo+" (ya usada en Conciliacion N° "+d.conciliacion_id+")</li>";}).join("");document.getElementById("modal-dup").classList.remove("hidden");}\nfunction cerrarModal(){document.getElementById("modal-dup").classList.add("hidden");}\nfunction confirmarDuplicados(){document.getElementById("modal-dup").classList.add("hidden");runConciliation(true);}\nfunction renderStats(data){document.getElementById("stats-grid").innerHTML="<div class=\\"stat\\"><div class=\\"stat-label\\">Conciliacion</div><div class=\\"stat-value stat-num\\">N° "+data.conciliacion_id+"</div></div><div class=\\"stat\\"><div class=\\"stat-label\\">Comprobantes ingresados</div><div class=\\"stat-value\\">"+data.total_comp+"</div></div><div class=\\"stat\\"><div class=\\"stat-label\\">Verificados</div><div class=\\"stat-value green\\">"+data.verificados+"</div></div><div class=\\"stat\\"><div class=\\"stat-label\\">Sin coincidencia</div><div class=\\"stat-value amber\\">"+(data.total_comp-data.verificados)+"</div></div>";}\nfunction renderCompResult(comps){document.getElementById("comp-body").innerHTML=comps.map(function(c){var chip=c.verificado?"<div class=\\"chip chip-v\\"><div class=\\"dot dot-g\\"></div>Verificado</div>":"<div class=\\"chip chip-r\\"><div class=\\"dot dot-r\\"></div>Sin coincidencia</div>";var monto=c.monto?"$"+c.monto.toLocaleString("es-AR",{minimumFractionDigits:0}):"-";return"<tr class=\\"+"+(c.verificado?"verified":"sinmatch")+"\\">"+"<td style=\\"font-size:12px\\">"+c.file+"</td>"+"<td style=\\"color:#6B7280\\">"+c.fecha+"</td>"+"<td style=\\"text-align:right;font-weight:600\\">"+monto+"</td>"+"<td>"+(c.nombre||"-")+"</td>"+"<td>"+chip+"</td>"+"<td><div class=\\"match-tag\\">"+(c.match||"-")+"</div></td>"+"</tr>";}).join("");document.getElementById("results-section").classList.remove("hidden");}\nfunction renderSinMatch(items){var s=document.getElementById("sinmatch-section");if(!items||items.length===0){s.classList.add("hidden");return;}document.getElementById("sinmatch-body").innerHTML=items.map(function(c){var monto=c.monto?"$"+c.monto.toLocaleString("es-AR",{minimumFractionDigits:0}):"-";return"<tr class=\\"sinmatch\\"><td style=\\"font-size:12px\\">"+c.file+"</td><td style=\\"color:#6B7280\\">"+c.fecha+"</td><td style=\\"text-align:right;font-weight:600\\">"+monto+"</td><td>"+(c.nombre||"-")+"</td><td>"+(c.banco_origen||"-")+"</td><td style=\\"font-family:monospace;font-size:11px\\">"+(c.cuit||"-")+"</td></tr>";}).join("");s.classList.remove("hidden");}\nfunction exportFile(fmt){if(!lastCompResult.length)return;var fd=new FormData();fd.append("rows",JSON.stringify(lastRows));fd.append("comp_result",JSON.stringify(lastCompResult));fd.append("sin_match",JSON.stringify(lastSinMatch));fd.append("format",fmt);fd.append("bank_label",lastBankLabel);fd.append("conciliacion_id",lastCid);fetch("/exportar",{method:"POST",body:fd}).then(function(r){return r.blob();}).then(function(blob){var url=URL.createObjectURL(blob);var a=document.createElement("a");a.href=url;a.download=fmt==="xlsx"?"conciliacion_"+lastCid+".xlsx":"conciliacion_"+lastCid+".pdf";a.click();URL.revokeObjectURL(url);});}\n</script>\n</body>\n</html>'

@app.route("/")
def index():
    return HTML_PAGE

@app.route("/conciliar",methods=["POST"])
def conciliar_route():
    if API_KEY.startswith("sk-ant-..."): return jsonify({"error":"API key no configurada. Agregar en Railway Variables: ANTHROPIC_API_KEY"})
    try:
        planilla=request.files["planilla"]; comps=request.files.getlist("comprobantes")
        bank_key=request.form.get("bank","Ciudad"); force=request.form.get("force","false")=="true"
        plan_path=os.path.join(UPLOAD_FOLDER,"planilla"+Path(planilla.filename).suffix); planilla.save(plan_path)
        comp_paths=[]
        for f in comps:
            p=os.path.join(UPLOAD_FOLDER,f.filename); f.save(p); comp_paths.append(p)
        if not force:
            dups=check_duplicados(comp_paths)
            if dups: return jsonify({"duplicados":dups})
        result=conciliar(plan_path,comp_paths,bank_key)
        comp_data=[{"file":c["file"],"hash":c["hash"],"monto":c.get("monto"),"fecha":c.get("fecha"),"nombre":c.get("nombre")} for c in result["comp_result"]]
        cid=guardar_conciliacion(bank_key,comp_data,result["verificados"])
        result["conciliacion_id"]=cid
        return jsonify(result)
    except Exception as e: return jsonify({"error":str(e)})

@app.route("/exportar",methods=["POST"])
def exportar_route():
    rows=json.loads(request.form["rows"]); comp_result=json.loads(request.form["comp_result"])
    sin_match=json.loads(request.form.get("sin_match","[]")); fmt=request.form["format"]
    bank_label=request.form.get("bank_label",""); cid=request.form.get("conciliacion_id","")
    if fmt=="xlsx":
        buf=export_excel(rows,comp_result,sin_match,cid,bank_label)
        return send_file(buf,download_name=f"conciliacion_{cid}.xlsx",as_attachment=True,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        buf=export_pdf(comp_result,cid,bank_label)
        return send_file(buf,download_name=f"conciliacion_{cid}.pdf",as_attachment=True,mimetype="application/pdf")

if __name__=="__main__":
    init_db()
    port=int(os.environ.get("PORT",5000))
    host="0.0.0.0" if os.environ.get("RAILWAY_ENVIRONMENT") else "127.0.0.1"
    if host=="127.0.0.1":
        print("="*50); print("  Conciliador Bancario"); print("  http://localhost:5000"); print("="*50)
    app.run(debug=False,host=host,port=port)
init_db()
