"""
Conciliador Bancario Web
========================
Correr local: python app.py  →  http://localhost:5000
Railway:      gunicorn app:app --timeout 120
"""

import os, re, json, base64, io, sqlite3, hashlib, time, logging
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
EXTRACT_WORKERS = int(os.environ.get("EXTRACT_WORKERS", "5"))   # lecturas de IA en paralelo
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("conciliador")

BANK_CONFIG = {
    "Ciudad": {"label":"Banco Ciudad (Argentina)","filter_starts":["TRANSFER","PAGO","DEPOSITO"],"exclude_cuits":["30708635754","30711747709"],"moneda":"ARS"},
    "ITAU":   {"label":"ITAU (Paraguay)",          "filter_starts":["TRANSFER","PAGO","TRANSFERENCIA"],"exclude_cuits":[],"moneda":"PYG"},
    "UENO":   {"label":"UENO (Paraguay)",           "filter_starts":["TRANSFER","PAGO","TRANSFERENCIA"],"exclude_cuits":[],"moneda":"PYG"},
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
            banco TEXT, moneda TEXT, total_comp INTEGER, verificados INTEGER,
            sin_match INTEGER, total_importe REAL, detalle TEXT,
            descripcion TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        -- Migracion: agregar columnas si no existen

        CREATE TABLE IF NOT EXISTS transferencias_verificadas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            conciliacion_id INTEGER, archivo TEXT, file_hash TEXT,
            monto REAL, fecha TEXT, cuit TEXT, nombre TEXT, banco_origen TEXT, referencia TEXT,
            mov_fecha TEXT, mov_importe REAL, mov_desc TEXT, criterio TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_tv_monto ON transferencias_verificadas(monto);
        CREATE TABLE IF NOT EXISTS comp_usados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_hash TEXT UNIQUE, nombre_archivo TEXT,
            monto REAL, fecha_comp TEXT, nombre_origen TEXT,
            conciliacion_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    conn.commit()
    # Migracion segura: agregar columnas nuevas si no existen
    for col, typedef in [("moneda","TEXT"), ("total_importe","REAL"), ("detalle","TEXT"), ("descripcion","TEXT")]:
        try: conn.execute(f"ALTER TABLE conciliaciones ADD COLUMN {col} {typedef}")
        except: pass
    conn.commit()
    # Backfill: cargar las conciliaciones viejas que todavia no estan en la base de transferencias
    try:
        hechas={r[0] for r in conn.execute("SELECT DISTINCT conciliacion_id FROM transferencias_verificadas")}
        for r in conn.execute("SELECT id,detalle FROM conciliaciones WHERE detalle IS NOT NULL").fetchall():
            if r["id"] in hechas: continue
            try: det=json.loads(r["detalle"] or "[]")
            except: continue
            _insertar_verificadas(conn,r["id"],det)
        conn.commit()
    except Exception as e: print("backfill:",e)
    conn.close()

def _norm_ref(r):
    r=re.sub(r"[^A-Za-z0-9]","",str(r or "")).upper()
    return r if len(r)>=6 else ""

def _insertar_verificadas(conn,cid,comp_result):
    for c in comp_result or []:
        if not c.get("verificado"): continue
        conn.execute("""INSERT INTO transferencias_verificadas
            (conciliacion_id,archivo,file_hash,monto,fecha,cuit,nombre,banco_origen,referencia,mov_fecha,mov_importe,mov_desc,criterio)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (cid,c.get("file"),c.get("hash"),c.get("monto"),c.get("fecha"),clean_cuit(c.get("cuit")),c.get("nombre"),
             c.get("banco_origen"),_norm_ref(c.get("referencia")),c.get("mov_fecha"),c.get("mov_importe"),c.get("mov_desc"),c.get("match")))

def cargar_verificadas():
    conn=get_db()
    rows=[dict(r) for r in conn.execute("SELECT * FROM transferencias_verificadas").fetchall()]
    conn.close(); return rows

def es_misma_transferencia(c,p):
    """c: comprobante actual (monto, fecha, cuit, nombre, referencia); p: registro guardado."""
    rc,rp=_norm_ref(c.get("referencia")),p.get("referencia") or ""
    if rc and rp: return rc==rp
    if not c.get("monto") or not p.get("monto") or abs(float(c["monto"])-float(p["monto"]))>0.5: return False
    if not c.get("fecha") or c.get("fecha")!=p.get("fecha"): return False
    cc,cp=clean_cuit(c.get("cuit"))[-10:],clean_cuit(p.get("cuit"))[-10:]
    if cc and cp: return cc==cp
    return len(name_tokens(c.get("nombre"))&name_tokens(p.get("nombre")))>=2

def mov_key(fecha,importe,desc):
    d=re.sub(r"\s+"," ",str(desc or "")).strip().upper()
    return f"{fecha}|{round(float(importe or 0),2)}|{d}"

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

def guardar_conciliacion(banco, comp_data_list, verificados, moneda="ARS", total_importe=0, detalle=None, descripcion=""):
    conn = get_db()
    cur = conn.execute("INSERT INTO conciliaciones (banco,moneda,total_comp,verificados,sin_match,total_importe,detalle,descripcion) VALUES (?,?,?,?,?,?,?,?)",
        (banco, moneda, len(comp_data_list), verificados, sum(1 for c in (detalle or []) if not c.get("verificado") and not c.get("duplicado")) if detalle else len(comp_data_list)-verificados, total_importe, json.dumps(detalle or []), descripcion or ""))
    cid = cur.lastrowid
    for item in comp_data_list:
        try:
            conn.execute("INSERT OR IGNORE INTO comp_usados (file_hash,nombre_archivo,monto,fecha_comp,nombre_origen,conciliacion_id) VALUES (?,?,?,?,?,?)",
                (item["hash"],item["file"],item.get("monto"),item.get("fecha"),item.get("nombre"),cid))
        except: pass
    _insertar_verificadas(conn,cid,detalle)
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
    try:
        if "," in s and "." in s:
            if s.rfind(",")>s.rfind("."): r=float(s.replace(".","").replace(",","."))
            else: r=float(s.replace(",",""))
        elif "," in s:
            r=float(s.replace(",","")) if s.count(",")>1 or len(s.split(",")[-1])==3 else float(s.replace(",","."))
        elif "." in s:
            r=float(s.replace(".","")) if s.count(".")>1 or len(s.split(".")[-1])==3 else float(s)
        else: r=float(s)
    except ValueError: return 0.0
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
    return {".pdf":"application/pdf",".jpg":"image/jpeg",".jpeg":"image/jpeg",".png":"image/png",".docx":"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}.get(Path(path).suffix.lower(),"")

def extract_images_from_docx(path):
    """Devuelve todas las imagenes del Word como lista de (bytes, content_type)."""
    from docx import Document
    doc = Document(path)
    return [(rel.target_part.blob, rel.target_part.content_type)
            for rel in doc.part.rels.values() if "image" in rel.reltype]

# IA
PROMPT_COMP = ("El archivo puede tener una o varias paginas/imagenes. Puede incluir resumenes de liquidacion, "
    "planillas u otros documentos que NO son comprobantes: ignoralos. "
    "Devolve SOLO un array JSON con un objeto por cada comprobante de transferencia o deposito bancario encontrado, con: "
    "fecha (DD/MM/YYYY), monto (numero sin simbolo ni separador de miles, punto como decimal), moneda, "
    "cuit_origen (solo digitos, del que ENVIA el dinero), nombre_origen, banco_origen, referencia. "
    "Null en los campos que no aparezcan. Si no hay ningun comprobante devolve []. Sin backticks ni texto extra.")

def extract_comprobantes(path):
    """Devuelve una lista de comprobantes encontrados en el archivo (puede ser vacia)."""
    client = anthropic.Anthropic(api_key=API_KEY, max_retries=4, timeout=120)
    suf = Path(path).suffix.lower()
    content = []
    if suf == ".docx":
        imgs = extract_images_from_docx(path)
        if not imgs: return []
        for data, ctype in imgs[:20]:
            content.append({"type":"image","source":{"type":"base64","media_type":ctype or "image/png",
                "data":base64.standard_b64encode(data).decode()}})
    else:
        is_img = suf in {".jpg",".jpeg",".png"}
        content.append({"type":"image" if is_img else "document",
            "source":{"type":"base64","media_type":get_media_type(path),"data":file_to_b64(path)}})
    content.append({"type":"text","text":PROMPT_COMP})
    resp = client.messages.create(model="claude-opus-4-5",max_tokens=2000,messages=[{"role":"user","content":content}])
    text = "".join(b.text for b in resp.content if hasattr(b,"text"))
    text = re.sub(r"```json|```","",text).strip()
    m = re.search(r"[\[{].*[\]}]", text, re.S)
    try: data = json.loads(m.group(0) if m else text)
    except Exception: return None
    if isinstance(data, dict): data = data.get("transferencias") or data.get("comprobantes") or [data]
    return [d for d in data if isinstance(d, dict)]

# Match
import unicodedata
NAME_STOP={"DE","DEL","LA","LAS","LOS","Y","SA","SAS","SRL","SAU","VAR","TRANSFER","TRANSFERENCIA","PAGO","DEPOSITO","PROVEEDORES","DATANET","INTERPYME"}
def name_tokens(t):
    t=unicodedata.normalize("NFKD",str(t or "")).encode("ascii","ignore").decode().upper()
    return {w for w in re.split(r"[^A-Z]+",t) if len(w)>=3 and w not in NAME_STOP}

def match_score(mov_fecha,mov_monto,mov_desc,comp):
    """Devuelve (score, razones, dias). Sin monto exacto no hay match."""
    c_monto=normalize_monto(comp.get("monto"))
    if not c_monto or not mov_monto: return 0,[],99
    if abs(mov_monto-c_monto)/max(mov_monto,c_monto)>=0.01: return 0,[],99
    score,razones=3,["monto exacto"]
    cuit_c=clean_cuit(comp.get("cuit_origen",""))
    cuit_d=cuit_en_desc(mov_desc)
    if cuit_c and cuit_d and cuit_c[-10:]==cuit_d[-10:]: score+=3; razones.append("CUIT")
    comunes=name_tokens(comp.get("nombre_origen",""))&name_tokens(mov_desc)
    if len(comunes)>=2: score+=2; razones.append("nombre")
    elif len(comunes)==1: score+=1; razones.append("nombre parcial")
    dias=99
    c_fecha=normalize_date(comp.get("fecha",""))
    if mov_fecha and c_fecha:
        try:
            dias=abs((date.fromisoformat(mov_fecha)-date.fromisoformat(c_fecha)).days)
            if dias==0: score+=2; razones.append("fecha exacta")
            elif dias<=3: score+=1; razones.append("fecha cercana")
        except: pass
    return score,razones,dias

def asignar(movs,comprobantes):
    """Asignacion 1 a 1: cada comprobante valida una sola fila y viceversa.
    Pasada 1: monto exacto + (CUIT/nombre/fecha), mejores puntajes primero.
    Pasada 2: solo monto exacto, fecha mas cercana (marcado para revisar)."""
    cands=[]
    for i,m in enumerate(movs):
        for j,c in enumerate(comprobantes):
            sc,rz,d=match_score(m["fecha"],m["monto"],m["desc"],c["data"])
            if sc>0: cands.append((sc,-d,i,j,rz,d))
    cands.sort(key=lambda x:(x[0],x[1]),reverse=True)
    row_to,comp_to={},{}
    for sc,_,i,j,rz,d in cands:
        if sc>=UMBRAL_MATCH and i not in row_to and j not in comp_to:
            row_to[i]=(j," + ".join(rz),False); comp_to[j]=i
    for sc,_,i,j,rz,d in sorted(cands,key=lambda x:x[5]):
        if i not in row_to and j not in comp_to and (d<=10 or d==99):
            row_to[i]=(j,"solo monto (revisar)",True); comp_to[j]=i
    return row_to,comp_to

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
    def leer(p):
        """Lee un archivo y devuelve la lista de comprobantes que contiene."""
        name=Path(p).name; fh=hash_file(p); out=[]
        try:
            items=extract_comprobantes(p)
            if not items or not any(normalize_monto(d.get("monto")) for d in items):
                items=extract_comprobantes(p) or items
        except Exception as e:
            log.exception("Error leyendo %s", name)
            return [{"file":name,"hash":fh,"data":{"monto":None,"fecha":None},"error":f"{type(e).__name__}: {e}"}]
        if not items:
            log.warning("Sin comprobantes en %s", name)
            return [{"file":name,"hash":fh,"data":{"monto":None,"fecha":None},"error":"No se encontro ningun comprobante en el archivo"}]
        for k,data in enumerate(items,1):
            f=normalize_date(data.get("fecha",""))
            if f and not f.startswith(("2025","2026","2027")): data["fecha"]=None
            out.append({"file":name if len(items)==1 else f"{name} ({k}/{len(items)})","hash":fh,"data":data,
                "monto":normalize_monto(data.get("monto")),"fecha":normalize_date(data.get("fecha","")),
                "nombre":data.get("nombre_origen",""),"banco_origen":data.get("banco_origen",""),"cuit":data.get("cuit_origen",""),
                "referencia":data.get("referencia") or ""})
        return out
    t0=time.time()
    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=EXTRACT_WORKERS) as ex:
        comprobantes=[c for lista in ex.map(leer,comp_paths) for c in lista]
    log.info("Lectura IA: %d archivos -> %d comprobantes, %d con error, %.0fs",
             len(comp_paths),len(comprobantes),sum(1 for c in comprobantes if c.get("error")),time.time()-t0)
    movs=[]
    for _,mov in ing.iterrows():
        movs.append({"fecha":normalize_date(mov.get(cols["fecha"],"") if cols["fecha"] else ""),
            "monto":normalize_monto(mov.get(cols["monto"],0)),
            "desc":str(mov.get(cols["desc"],"") if cols["desc"] else "").strip()})
    # Repetidos: contra la base de transferencias ya verificadas y dentro de esta misma carga
    prev=cargar_verificadas()
    hash_vistos={}; aceptados=[]
    for c in comprobantes:
        c["dup"]=None
        if not c.get("monto"): continue
        p=next((p for p in prev if es_misma_transferencia(c,p)),None)
        if p:
            c["dup"]=f"Ya verificada en Conciliacion N° {p['conciliacion_id']} ({p['archivo']})"; continue
        base=c["file"].split(" (")[0]
        if c["hash"] in hash_vistos and hash_vistos[c["hash"]]!=base:
            c["dup"]=f"Mismo archivo que {hash_vistos[c['hash']]} en esta carga"; continue
        hash_vistos.setdefault(c["hash"],base)
        q=next((q for q in aceptados if q["file"].split(" (")[0]!=base and
                es_misma_transferencia(c,{"monto":q["monto"],"fecha":q["fecha"],"cuit":q["cuit"],"nombre":q["nombre"],
                                          "referencia":_norm_ref(q.get("referencia"))})),None)
        if q: c["dup"]=f"Misma transferencia que {q['file']} en esta carga"; continue
        aceptados.append(c)
    # Filas del extracto ya conciliadas antes (respeta repeticiones legitimas del mismo movimiento)
    from collections import Counter
    cupo=Counter(); cid_mov={}
    for p in prev:
        if p.get("mov_desc"):
            k=mov_key(p["mov_fecha"],p["mov_importe"],p["mov_desc"]); cupo[k]+=1; cid_mov[k]=p["conciliacion_id"]
    for m in movs:
        k=mov_key(m["fecha"],m["monto"],m["desc"]); m["usado"]=None
        if cupo[k]>0: cupo[k]-=1; m["usado"]=cid_mov[k]
    activos_c=[j for j,c in enumerate(comprobantes) if not c["dup"]]
    activos_m=[i for i,m in enumerate(movs) if not m["usado"]]
    rt,ct=asignar([movs[i] for i in activos_m],[comprobantes[j] for j in activos_c])
    row_to={activos_m[i]:(activos_c[a[0]],a[1],a[2]) for i,a in rt.items()}
    comp_to={activos_c[j]:activos_m[i] for j,i in ct.items()}
    rows=[]
    for i,m in enumerate(movs):
        a=row_to.get(i)
        rows.append({"fecha":m["fecha"],"descripcion":m["desc"],"importe":round(m["monto"],2),
            "estado":(f"Ya conciliado (N° {m['usado']})" if m.get("usado") else
                      ("Verificado (monto)" if a and a[2] else "Verificado") if a else "Sin comprobante"),
            "comprobante":comprobantes[a[0]]["file"] if a else "","match":a[1] if a else ""})
    comp_result=[]
    for j,c in enumerate(comprobantes):
        i=comp_to.get(j); a=row_to.get(i) if i is not None else None
        comp_result.append({"file":c["file"],"hash":c["hash"],"fecha":c.get("fecha",""),"monto":c.get("monto",0),
            "nombre":c.get("nombre","—"),"banco_origen":c.get("banco_origen","—"),"cuit":c.get("cuit","—"),
            "verificado":a is not None,"revisar":bool(a and a[2]),"match":a[1] if a else "",
            "mov_desc":movs[i]["desc"] if i is not None else "",
            "mov_fecha":movs[i]["fecha"] if i is not None else "",
            "mov_importe":movs[i]["monto"] if i is not None else None,
            "referencia":c.get("referencia",""),"duplicado":c.get("dup") or "",
            "error_extraccion":c.get("error","") or ("" if c.get("monto") else "No se pudo leer el monto")})
    sin_match=[c for c in comp_result if not c["verificado"] and not c["duplicado"]]
    duplicados=[c for c in comp_result if c["duplicado"]]
    return {"rows":rows,"comp_result":comp_result,"sin_match":sin_match,"duplicados":duplicados,
            "total_comp":len(comprobantes),"verificados":sum(1 for c in comp_result if c["verificado"])}

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
            "Verificado" if c["verificado"] else ("Repetido: "+c["duplicado"] if c.get("duplicado") else "Sin coincidencia"),c["match"]])
        fill=verde if c["verificado"] else (amarillo if c.get("duplicado") else rojo)
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

def export_pdf(comp_result,cid,bank_label,descripcion=""):
    buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=landscape(A4),leftMargin=1.5*cm,rightMargin=1.5*cm,topMargin=2*cm,bottomMargin=1.5*cm)
    styles=getSampleStyleSheet()
    ts=ParagraphStyle("t",parent=styles["Heading1"],fontSize=14,textColor=colors.HexColor("#1F4E79"),spaceAfter=4)
    ss=ParagraphStyle("s",parent=styles["Normal"],fontSize=9,textColor=colors.grey,spaceAfter=4)
    ds=ParagraphStyle("d",parent=styles["Normal"],fontSize=10,textColor=colors.HexColor("#374151"),spaceAfter=14)
    cs=ParagraphStyle("c",parent=styles["Normal"],fontSize=8,leading=11)
    total=len(comp_result); verif=sum(1 for c in comp_result if c.get("verificado"))
    story=[
        Paragraph(f"Conciliacion N{chr(176)} {cid} - {bank_label}",ts),
        Paragraph(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Ingresados: {total}  |  Verificados: {verif}  |  Sin coincidencia: {total-verif}",ss),
    ]
    if descripcion:
        story.append(Paragraph(f"Descripcion: {descripcion}",ds))
    verif_list=[c for c in comp_result if c.get("verificado")]
    if verif_list:
        story.append(Paragraph("Comprobantes verificados",ParagraphStyle("h2",parent=styles["Heading2"],fontSize=11,textColor=colors.HexColor("#166534"),spaceAfter=6)))
        td=[["Archivo","Fecha","Monto","Remitente","Criterios de match"]]
        for c in verif_list:
            td.append([Paragraph(c.get("file",""),cs),c.get("fecha") or "-",f"${c['monto']:,.0f}" if c.get("monto") else "-",Paragraph(c.get("nombre") or "-",cs),c.get("match") or "-"])
        t=Table(td,colWidths=[5*cm,2.5*cm,3*cm,7*cm,8*cm],repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#166534")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("FONTNAME",(0,1),(-1,-1),"Helvetica"),
            ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#C6EFCE")),("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
            *[("BACKGROUND",(0,k+1),(-1,k+1),colors.HexColor("#FFEB9C")) for k,c in enumerate(verif_list) if c.get("revisar")],
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),5),
        ])); story.append(t)
    dup_list=[c for c in comp_result if c.get("duplicado")]
    sin_list=[c for c in comp_result if not c.get("verificado") and not c.get("duplicado")]
    if sin_list:
        story.append(Spacer(1,0.7*cm))
        story.append(Paragraph("Comprobantes sin coincidencia",ParagraphStyle("h2b",parent=styles["Heading2"],fontSize=11,textColor=colors.HexColor("#991B1B"),spaceAfter=6)))
        td2=[["Archivo","Fecha","Monto","Remitente","Banco origen"]]
        for c in sin_list:
            td2.append([Paragraph(c.get("file",""),cs),c.get("fecha") or "-",f"${c['monto']:,.0f}" if c.get("monto") else "-",Paragraph(c.get("nombre") or "-",cs),c.get("banco_origen") or "-"])
        t2=Table(td2,colWidths=[5*cm,2.5*cm,3*cm,7*cm,8*cm],repeatRows=1)
        t2.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#C0392B")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("FONTNAME",(0,1),(-1,-1),"Helvetica"),
            ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#FFCCCC")),("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),5),
        ])); story.append(t2)
    if dup_list:
        story.append(Spacer(1,0.7*cm))
        story.append(Paragraph("Comprobantes repetidos (no se tomaron)",ParagraphStyle("h2c",parent=styles["Heading2"],fontSize=11,textColor=colors.HexColor("#5B21B6"),spaceAfter=6)))
        td3=[["Archivo","Fecha","Monto","Remitente","Motivo"]]
        for c in dup_list:
            td3.append([Paragraph(c.get("file",""),cs),c.get("fecha") or "-",f"${c['monto']:,.0f}" if c.get("monto") else "-",Paragraph(c.get("nombre") or "-",cs),Paragraph(c.get("duplicado",""),cs)])
        t3=Table(td3,colWidths=[5*cm,2.5*cm,3*cm,6*cm,9*cm],repeatRows=1)
        t3.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#5B21B6")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),
            ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#EDE9FE")),("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ])); story.append(t3)
    doc.build(story); buf.seek(0); return buf

HTML_PAGE = '<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width,initial-scale=1">\n<title>Conciliador — La Santaniana</title>\n<link rel="preconnect" href="https://fonts.googleapis.com">\n<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">\n<style>\n:root{--red:#D4141C;--blue:#1B3FA0;--blue-dark:#152F7A;--red-light:#FFF0F0;--blue-light:#EEF2FF;--gray-50:#F8F9FB;--gray-100:#F1F3F6;--gray-200:#E4E7ED;--gray-400:#9CA3AF;--gray-600:#4B5563;--gray-800:#1F2937;--green:#16A34A;--green-light:#DCFCE7;--amber:#92400E;--amber-light:#FEF3C7}\n*{box-sizing:border-box;margin:0;padding:0}\nbody{font-family:"Inter",Arial,sans-serif;background:var(--gray-50);color:var(--gray-800);font-size:14px;min-height:100vh}\n/* HEADER */\nheader{background:#fff;padding:0;box-shadow:0 1px 8px rgba(0,0,0,.06);border-bottom:3px solid transparent;border-image:linear-gradient(90deg,var(--red) 0%,var(--red) 30%,var(--blue) 30%,var(--blue) 100%) 1}\n.header-inner{max-width:1060px;margin:0 auto;padding:.9rem 1.5rem;display:flex;align-items:center;justify-content:space-between}\n.header-logo{height:38px;object-fit:contain}\n.header-right{display:flex;align-items:center;gap:8px;color:rgba(255,255,255,.75);font-size:12px;font-weight:500}\n.header-pill{background:var(--blue-light);border-radius:20px;padding:5px 14px;font-size:12px;font-weight:700;color:var(--blue);letter-spacing:.03em}\n/* LAYOUT */\n.container{max-width:1060px;margin:0 auto;padding:1.75rem 1.5rem}\n/* CARDS */\n.card{background:#fff;border:1px solid var(--gray-200);border-radius:12px;padding:1.4rem 1.5rem;margin-bottom:1.1rem;box-shadow:0 1px 4px rgba(0,0,0,.04)}\n.card-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--gray-400);margin-bottom:1rem;display:flex;align-items:center;gap:6px}\n.card-title::before{content:"";display:block;width:3px;height:13px;border-radius:2px;background:var(--blue)}\n/* FORM */\n.form-row{display:grid;grid-template-columns:280px 1fr;gap:1rem;align-items:end;margin-bottom:0}\n.label{font-size:12px;font-weight:600;color:var(--gray-600);margin-bottom:5px}\nselect,input[type=text]{width:100%;padding:.6rem .85rem;border:1.5px solid var(--gray-200);border-radius:8px;font-size:13.5px;font-family:inherit;outline:none;transition:.15s;background:#fff;color:var(--gray-800)}\nselect:focus,input[type=text]:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(27,63,160,.1)}\n/* UPLOAD */\n.upload-grid{display:grid;grid-template-columns:1fr 1fr;gap:1rem}\n.drop{border:2px dashed var(--gray-200);border-radius:10px;padding:1.4rem 1rem;text-align:center;cursor:pointer;transition:.18s;user-select:none;background:var(--gray-50)}\n.drop:hover,.drop.over{background:var(--blue-light);border-color:var(--blue);transform:translateY(-1px)}\n.drop-icon{font-size:28px;margin-bottom:6px;opacity:.7}\n.drop-title{font-weight:600;font-size:13.5px;color:var(--gray-800);margin-bottom:3px}\n.drop-sub{font-size:12px;color:var(--gray-400)}\n.file-list{margin-top:.65rem;display:flex;flex-direction:column;gap:5px}\n.file-item{display:flex;align-items:center;gap:8px;font-size:12px;padding:5px 10px;background:var(--gray-50);border-radius:7px;border:1px solid var(--gray-200)}\n.badge{font-size:10px;font-weight:700;padding:2px 7px;border-radius:4px;letter-spacing:.02em}\n.badge-xlsx{background:#DCFCE7;color:#15803D}\n.badge-pdf{background:#FEE2E2;color:#B91C1C}\n.badge-img{background:#DBEAFE;color:#1D4ED8}\n.badge-docx{background:#EDE9FE;color:#6D28D9}\n.fname{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--gray-600)}\n.rm{cursor:pointer;color:var(--gray-400);font-size:16px;line-height:1;padding:0 2px;transition:.1s}\n.rm:hover{color:var(--red)}\n/* BUTTONS */\n.btn-row{display:flex;gap:10px;margin-top:.1rem}\n.btn-primary{flex:1;padding:.8rem 1.2rem;font-size:14.5px;font-weight:600;font-family:inherit;background:linear-gradient(135deg,var(--blue) 0%,var(--blue-dark) 100%);color:#fff;border:none;border-radius:9px;cursor:pointer;transition:.18s;box-shadow:0 2px 8px rgba(27,63,160,.3);letter-spacing:.01em}\n.btn-primary:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 4px 14px rgba(27,63,160,.4)}\n.btn-primary:disabled{opacity:.45;cursor:not-allowed;transform:none;box-shadow:none}\n.btn-secondary{padding:.8rem 1.2rem;font-size:13.5px;font-weight:600;font-family:inherit;background:#fff;color:var(--gray-600);border:1.5px solid var(--gray-200);border-radius:9px;cursor:pointer;transition:.18s;white-space:nowrap}\n.btn-secondary:hover{background:var(--gray-100);border-color:var(--gray-400)}\n/* PROGRESS */\n.progress-wrap{margin-top:.85rem}\n.progress-bar{height:4px;background:var(--gray-200);border-radius:2px;overflow:hidden}\n.progress-fill{height:100%;background:linear-gradient(90deg,var(--blue),var(--red));border-radius:2px;transition:width .5s ease}\n.status-msg{font-size:12.5px;color:var(--gray-400);text-align:center;margin-top:.5rem;min-height:18px}\n.hidden{display:none!important}\n/* STATS */\n.stats-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:1.1rem}\n.stat{background:#fff;border:1px solid var(--gray-200);border-radius:10px;padding:1rem 1.1rem;box-shadow:0 1px 3px rgba(0,0,0,.04)}\n.stat-label{font-size:11px;color:var(--gray-400);font-weight:600;text-transform:uppercase;letter-spacing:.05em;margin-bottom:5px}\n.stat-value{font-size:24px;font-weight:700;color:var(--gray-800)}\n.stat-num{font-size:15px;color:var(--blue);font-weight:700}\n.c-green{color:var(--green)}.c-red{color:var(--red)}\n/* RESULTS */\n.results-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:.85rem;flex-wrap:wrap;gap:.5rem}\n.export-row{display:flex;gap:7px;flex-wrap:wrap}\n.btn-exp{padding:.4rem .9rem;font-size:12.5px;font-weight:600;font-family:inherit;border-radius:7px;cursor:pointer;transition:.15s;border:1.5px solid}\n.btn-exp:hover{transform:translateY(-1px)}\n.btn-xlsx{border-color:#15803D;color:#15803D;background:#fff}\n.btn-xlsx:hover{background:#DCFCE7}\n.btn-pdf{border-color:#B91C1C;color:#B91C1C;background:#fff}\n.btn-pdf:hover{background:#FEE2E2}\n.btn-cerrar{border-color:var(--green);color:#fff;background:var(--green);font-weight:700}\n.btn-cerrar:hover:not(:disabled){background:#15803D}\n.btn-cerrar:disabled{opacity:.45;cursor:not-allowed}\n/* TABLE */\ntable{width:100%;border-collapse:collapse;font-size:13px}\nth{text-align:left;padding:9px 11px;font-size:10.5px;font-weight:700;color:var(--gray-400);border-bottom:2px solid var(--gray-200);white-space:nowrap;text-transform:uppercase;letter-spacing:.05em}\ntd{padding:9px 11px;border-bottom:1px solid var(--gray-100);vertical-align:middle}\ntr:last-child td{border-bottom:none}\n#comp-body td:nth-child(2),#dup-body td:nth-child(2),#sinmatch-body td:nth-child(2){white-space:nowrap}\ntr.verified td{background:#F0FDF4}\ntr.sinmatch td{background:#FFF5F5}\ntr:hover td{filter:brightness(.97)}\n.chip{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px}\n.chip-v{background:var(--green-light);color:var(--green)}\n.chip-r{background:#FEE2E2;color:#B91C1C}\n.dot{width:6px;height:6px;border-radius:50%;flex-shrink:0}\n.chip-a{background:var(--amber-light);color:var(--amber)}\n.dot-a{background:#D97706}\ntr.revisar td{background:#FFFBEB}\n.dot-g{background:var(--green)}.dot-r{background:var(--red)}\n.match-tag{font-size:11px;color:var(--gray-400)}\n/* SIN MATCH */\n.sinmatch-card{background:#fff;border:1px solid #FECACA;border-radius:12px;padding:1.4rem 1.5rem;margin-bottom:1.1rem;box-shadow:0 1px 4px rgba(212,20,28,.06)}\n.sinmatch-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--red);margin-bottom:1rem;display:flex;align-items:center;gap:6px}\n.sinmatch-title::before{content:"";display:block;width:3px;height:13px;border-radius:2px;background:var(--red)}\n/* ERROR */\n.error-box{background:#FEF2F2;border:1px solid #FECACA;border-radius:9px;padding:.9rem 1.1rem;margin-top:.75rem;font-size:13px;color:#B91C1C;font-weight:500}\n.scroll-wrap{overflow-x:auto}\n/* MODAL */\n.modal-overlay{position:fixed;inset:0;background:rgba(15,23,42,.55);backdrop-filter:blur(3px);display:flex;align-items:center;justify-content:center;z-index:1000}\n.modal{background:#fff;border-radius:14px;padding:2rem;max-width:520px;width:90%;box-shadow:0 12px 40px rgba(0,0,0,.2)}\n.modal h2{font-size:15px;font-weight:700;color:#B91C1C;margin-bottom:.4rem}\n.modal p{font-size:13px;color:var(--gray-600);margin-bottom:.9rem;line-height:1.5}\n.modal-list{background:#FEF2F2;border-radius:8px;padding:.75rem 1rem;margin-bottom:1.1rem;font-size:13px}\n.modal-list li{padding:3px 0;color:var(--gray-800);list-style:none}\n.modal-btns{display:flex;gap:.75rem;justify-content:flex-end}\n.btn-cancel{padding:.55rem 1.15rem;border:1.5px solid var(--gray-200);border-radius:8px;background:#fff;cursor:pointer;font-size:13px;font-family:inherit;font-weight:500}\n.btn-force{padding:.55rem 1.15rem;border:none;border-radius:8px;background:#B91C1C;color:#fff;cursor:pointer;font-size:13px;font-family:inherit;font-weight:600}\n/* HISTORIAL */\n.hist-btn{background:#fff;border:1.5px solid var(--gray-200);padding:.55rem 1.2rem;border-radius:9px;cursor:pointer;font-size:13px;font-family:inherit;font-weight:600;color:var(--gray-600);display:inline-flex;align-items:center;gap:6px;transition:.15s}\n.chip-d{background:#EDE9FE;color:#5B21B6}\n.dot-d{background:#7C3AED}\ntr.dup td{background:#F5F3FF}\n.dup-card{background:#fff;border:1px solid #DDD6FE;border-radius:12px;padding:1.4rem 1.5rem;margin-bottom:1.1rem}\n.dup-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#5B21B6;margin-bottom:1rem}\n.search-in{width:100%;max-width:420px;margin-bottom:.85rem}\n.btn-del{padding:4px 9px;font-size:12px;border:1.5px solid var(--gray-200);color:var(--gray-600);background:#fff;border-radius:6px;cursor:pointer;font-family:inherit}\n.btn-del:hover{border-color:var(--red);color:var(--red)}\n@media(max-width:720px){.stats-grid{grid-template-columns:repeat(2,1fr)}.upload-grid,.form-row{grid-template-columns:1fr}}\n.hist-btn:hover{background:var(--gray-50);border-color:var(--blue);color:var(--blue)}\n</style>\n</head>\n<body>\n<div class="modal-overlay hidden" id="modal-dup">\n  <div class="modal">\n    <h2>&#x26A0;&#xFE0F; Comprobantes ya utilizados</h2>\n    <p>Los siguientes comprobantes ya fueron procesados en una conciliacion anterior:</p>\n    <ul class="modal-list" id="modal-list"></ul>\n    <p>&#xBF;Deseas incluirlos de todas formas en esta conciliacion?</p>\n    <div class="modal-btns">\n      <button class="btn-cancel" onclick="cerrarModal()">Cancelar</button>\n      <button class="btn-force" onclick="confirmarDuplicados()">S&#xED;, incluir igual</button>\n    </div>\n  </div>\n</div>\n<header>\n  <div class="header-inner">\n    <img src="data:image/jpeg;base64,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" class="header-logo" alt="La Santaniana">\n    <div class="header-right">\n      <span class="header-pill">Conciliador Bancario</span>\n    </div>\n  </div>\n</header>\n<div class="container">\n  <div class="card">\n    <div class="card-title">Configuraci&#xF3;n</div>\n    <div class="form-row">\n      <div>\n        <div class="label">Banco</div>\n        <select id="bank-sel">\n          <option value="Ciudad">Banco Ciudad (Argentina)</option>\n          <option value="ITAU">ITAU (Paraguay)</option>\n          <option value="UENO">UENO (Paraguay)</option>\n        </select>\n      </div>\n      <div>\n        <div class="label">Descripci&#xF3;n (opcional)</div>\n        <input type="text" id="inp-desc" placeholder="Ej: Agencias octubre 2026, semana 1">\n      </div>\n    </div>\n  </div>\n  <div class="card">\n    <div class="card-title">Archivos</div>\n    <div class="upload-grid">\n      <div>\n        <div class="drop" id="drop-xlsx">\n          <div class="drop-icon">&#x1F4CA;</div>\n          <div class="drop-title">Planilla de movimientos</div>\n          <div class="drop-sub">Click o arrastr&#xE1; el Excel del banco</div>\n        </div>\n        <input type="file" id="inp-xlsx" accept=".xlsx,.xls,.csv" style="display:none">\n        <div class="file-list" id="list-xlsx"></div>\n      </div>\n      <div>\n        <div class="drop" id="drop-comp">\n          <div class="drop-icon">&#x1F5C2;&#xFE0F;</div>\n          <div class="drop-title">Comprobantes de transferencia</div>\n          <div class="drop-sub">PDF, JPG, PNG o Word (.docx)</div>\n        </div>\n        <input type="file" id="inp-comp" accept=".pdf,.jpg,.jpeg,.png,.docx" multiple style="display:none">\n        <div class="file-list" id="list-comp"></div>\n      </div>\n    </div>\n  </div>\n  <div class="btn-row">\n    <button class="btn-primary" id="run-btn" disabled onclick="runConciliation(false)">&#x2728; Conciliar transferencias</button>\n    <button class="btn-secondary" id="nueva-btn" onclick="nuevaConciliacion()">&#x21BA; Nueva</button>\n  </div>\n  <div class="progress-wrap hidden" id="prog-wrap">\n    <div class="progress-bar"><div class="progress-fill" id="prog-fill" style="width:0%"></div></div>\n    <div class="status-msg" id="status-msg"></div>\n  </div>\n  <div id="error-area"></div>\n  <div id="results-section" class="hidden" style="margin-top:1.5rem">\n    <div class="stats-grid" id="stats-grid"></div>\n    <div class="card">\n      <div class="results-header">\n        <div class="card-title" style="margin-bottom:0">Comprobantes</div>\n        <div class="export-row">\n          <button class="btn-exp btn-xlsx" onclick="exportFile(\'xlsx\')">&#x2B07; Excel</button>\n          <button class="btn-exp btn-pdf"  onclick="exportFile(\'pdf\')">&#x2B07; PDF</button>\n          <button class="btn-exp btn-cerrar" id="cerrar-btn" onclick="cerrarConciliacion()" disabled>&#x2714; Cerrar conciliaci&#xF3;n</button>\n        </div>\n      </div>\n      <div class="scroll-wrap">\n        <table><thead><tr>\n          <th>Archivo</th><th style="width:90px">Fecha</th>\n          <th style="width:115px;text-align:right">Monto</th>\n          <th>Remitente</th><th style="width:150px">Estado</th><th>Criterios</th>\n        </tr></thead><tbody id="comp-body"></tbody></table>\n      </div>\n    </div>\n    <div id="dup-section" class="hidden">\n      <div class="dup-card">\n        <div class="dup-title">&#x1F501; Comprobantes repetidos &mdash; no se tomaron</div>\n        <div class="scroll-wrap">\n          <table><thead><tr>\n            <th>Archivo</th><th style="width:90px">Fecha</th>\n            <th style="width:115px;text-align:right">Monto</th>\n            <th>Remitente</th><th>Motivo</th>\n          </tr></thead><tbody id="dup-body"></tbody></table>\n        </div>\n      </div>\n    </div>\n    <div id="sinmatch-section" class="hidden">\n      <div class="sinmatch-card">\n        <div class="sinmatch-title">&#x26A0; Comprobantes sin coincidencia en el extracto</div>\n        <div class="scroll-wrap">\n          <table><thead><tr>\n            <th>Archivo</th><th style="width:90px">Fecha</th>\n            <th style="width:115px;text-align:right">Monto</th>\n            <th>Remitente</th><th>Banco</th><th>CUIT</th>\n          </tr></thead><tbody id="sinmatch-body"></tbody></table>\n        </div>\n      </div>\n    </div>\n  </div>\n  <div style="margin-top:2rem;margin-bottom:.5rem">\n    <button class="hist-btn" onclick="cargarHistorial()">&#x1F4CB; <span id="hist-btn-txt">Ver historial</span></button>\n    <button class="hist-btn" onclick="toggleTransf()" style="margin-left:6px">&#x1F50E; <span id="tv-btn-txt">Transferencias verificadas</span></button>\n  </div>\n  <div id="tv-section" class="hidden" style="margin-top:.75rem">\n    <div class="card">\n      <div class="card-title">Transferencias verificadas <span id="tv-total" style="text-transform:none;letter-spacing:0;font-weight:500"></span></div>\n      <input type="text" id="tv-q" class="search-in" placeholder="Buscar por nombre, CUIT, monto, archivo u operaci&oacute;n..." oninput="buscarTransfDebounced()">\n      <div class="scroll-wrap">\n        <table><thead><tr>\n          <th>Fecha</th><th style="text-align:right">Monto</th><th>Remitente</th><th>CUIT</th><th>Archivo</th><th>Conciliaci&oacute;n</th><th>Movimiento del extracto</th>\n        </tr></thead><tbody id="tv-body"></tbody></table>\n      </div>\n    </div>\n  </div>\n  <div id="hist-section" class="hidden" style="margin-top:.75rem">\n    <div class="card">\n      <div class="card-title">Historial de conciliaciones</div>\n      <div class="scroll-wrap">\n        <table><thead><tr>\n          <th>N&#xB0;</th><th>Fecha</th><th>Banco</th><th>Ingresados</th><th>Verificados</th><th>Sin coinc.</th><th>Total verificado</th><th style="min-width:280px">Descripci&#xF3;n / PDF</th>\n        </tr></thead><tbody id="hist-body"></tbody></table>\n      </div>\n    </div>\n  </div>\n</div>\n<script>\nvar procTimer=null,lastNDup=0,tvTimer=null,xlsxFile=null,compFiles=[],lastCompResult=[],lastRows=[],lastSinMatch=[],lastCid=null,lastBankLabel="",lastBankKey="",lastMoneda="",lastDesc="",conciliacionCerrada=false;\nfunction ext(n){return(n||"").split(".").pop().toLowerCase();}\nfunction badgeClass(n){var e=ext(n);if(e==="xlsx"||e==="xls"||e==="csv")return"badge-xlsx";if(e==="pdf")return"badge-pdf";if(e==="docx")return"badge-docx";return"badge-img";}\nfunction badgeLabel(n){var e=ext(n);if(e==="xlsx"||e==="xls")return"XLSX";if(e==="csv")return"CSV";if(e==="pdf")return"PDF";if(e==="docx")return"DOCX";return"IMG";}\nfunction renderXlsx(){var el=document.getElementById("list-xlsx");if(!xlsxFile){el.innerHTML="";return;}el.innerHTML="<div class=\\"file-item\\"><span class=\\"badge "+badgeClass(xlsxFile.name)+"\\">"+badgeLabel(xlsxFile.name)+"</span><span class=\\"fname\\">"+xlsxFile.name+"</span><span class=\\"rm\\" onclick=\\"xlsxFile=null;renderXlsx();check()\\">&#x2715;</span></div>";}\nfunction renderComp(){var el=document.getElementById("list-comp");el.innerHTML=compFiles.map(function(f,i){return"<div class=\\"file-item\\"><span class=\\"badge "+badgeClass(f.name)+"\\">"+badgeLabel(f.name)+"</span><span class=\\"fname\\">"+f.name+"</span><span class=\\"rm\\" onclick=\\"compFiles.splice("+i+",1);renderComp();check()\\">&#x2715;</span></div>";}).join("");}\nfunction check(){document.getElementById("run-btn").disabled=!(xlsxFile&&compFiles.length>0);}\ndocument.getElementById("drop-xlsx").addEventListener("click",function(){document.getElementById("inp-xlsx").click();});\ndocument.getElementById("drop-comp").addEventListener("click",function(){document.getElementById("inp-comp").click();});\ndocument.getElementById("inp-xlsx").addEventListener("change",function(){if(this.files&&this.files.length>0){xlsxFile=this.files[0];renderXlsx();check();}});\ndocument.getElementById("inp-comp").addEventListener("change",function(){if(this.files&&this.files.length>0){for(var i=0;i<this.files.length;i++)compFiles.push(this.files[i]);renderComp();check();}});\nfunction setupDrop(id,isXlsx){var el=document.getElementById(id);el.addEventListener("dragover",function(e){e.preventDefault();e.stopPropagation();el.classList.add("over");});el.addEventListener("dragleave",function(e){e.stopPropagation();el.classList.remove("over");});el.addEventListener("drop",function(e){e.preventDefault();e.stopPropagation();el.classList.remove("over");var files=Array.from(e.dataTransfer.files);if(isXlsx){var f=files.find(function(f){return["xlsx","xls","csv"].indexOf(ext(f.name))>=0;});if(f){xlsxFile=f;renderXlsx();check();}}else{var valid=files.filter(function(f){return["pdf","jpg","jpeg","png","docx"].indexOf(ext(f.name))>=0;});valid.forEach(function(f){compFiles.push(f);});if(valid.length){renderComp();check();}}});}\nsetupDrop("drop-xlsx",true);setupDrop("drop-comp",false);\nfunction setProgress(pct,msg){document.getElementById("prog-fill").style.width=pct+"%";document.getElementById("status-msg").textContent=msg;}\nfunction showError(msg){document.getElementById("error-area").innerHTML="<div class=\\"error-box\\">"+msg+"</div>";document.getElementById("prog-wrap").classList.add("hidden");document.getElementById("run-btn").disabled=false;}\nasync function runConciliation(force){\n  document.getElementById("error-area").innerHTML="";\n  document.getElementById("results-section").classList.add("hidden");\n  document.getElementById("prog-wrap").classList.remove("hidden");\n  document.getElementById("run-btn").disabled=true;\n  setProgress(5,"Subiendo archivos...");\n  var bankSel=document.getElementById("bank-sel");\n  lastBankLabel=bankSel.options[bankSel.selectedIndex].text;\n  var fd=new FormData();\n  fd.append("planilla",xlsxFile);\n  for(var i=0;i<compFiles.length;i++)fd.append("comprobantes",compFiles[i]);\n  fd.append("bank",bankSel.value);\n  fd.append("descripcion",document.getElementById("inp-desc").value||"");\n  fd.append("force",force?"true":"false");\n  try{\n    var t0=Date.now(),n=compFiles.length;\n    procTimer=setInterval(function(){var s=Math.round((Date.now()-t0)/1000);\n      setProgress(Math.min(15+s*1.2,92),"Leyendo "+n+" comprobantes con IA\\u2026 "+Math.floor(s/60)+":"+String(s%60).padStart(2,"0")+" \\u2014 no recargues la p\\u00e1gina");},1000);\n    window.onbeforeunload=function(){return "La conciliaci\\u00f3n sigue procesando.";};\n    var resp=await fetch("/conciliar",{method:"POST",body:fd});\n    var data=await resp.json();\n    if(data.error){showError(data.error);return;}\n    setProgress(100,"Listo. Revis\\u00e1 los resultados y cerr\\u00e1 cuando est\\u00e9s conforme.");\n    lastRows=data.rows;lastCompResult=data.comp_result;lastSinMatch=data.sin_match||[];lastCid=null;\n    lastBankKey=data.bank_key||bankSel.value;lastMoneda=data.moneda||"ARS";lastDesc=document.getElementById("inp-desc").value||"";\n    conciliacionCerrada=false;\n    var cb=document.getElementById("cerrar-btn");cb.disabled=false;cb.textContent="\\u2714 Cerrar conciliaci\\u00f3n";cb.style.background="";\n    data.n_dup=(data.duplicados||[]).length;lastNDup=data.n_dup;renderStats(data);renderCompResult(data.comp_result);renderSinMatch(data.sin_match||[]);renderDup(data.duplicados||[]);\n  }catch(e){showError("Error: "+e.message);}\n  finally{clearInterval(procTimer);window.onbeforeunload=null;document.getElementById("run-btn").disabled=false;}\n}\nfunction mostrarModal(dups){var lista=document.getElementById("modal-list");lista.innerHTML=dups.map(function(d){return"<li>\\u2022 "+d.archivo+" (Conciliaci\\u00f3n N\\u00b0 "+d.conciliacion_id+")</li>";}).join("");document.getElementById("modal-dup").classList.remove("hidden");}\nfunction cerrarModal(){document.getElementById("modal-dup").classList.add("hidden");}\nfunction confirmarDuplicados(){document.getElementById("modal-dup").classList.add("hidden");runConciliation(true);}\nfunction renderStats(data){var cidLabel=data.conciliacion_id?"N\\u00b0 "+data.conciliacion_id:"Pendiente";document.getElementById("stats-grid").innerHTML="<div class=\\"stat\\"><div class=\\"stat-label\\">Conciliaci\\u00f3n</div><div class=\\"stat-value stat-num\\">"+cidLabel+"</div></div><div class=\\"stat\\"><div class=\\"stat-label\\">Ingresados</div><div class=\\"stat-value\\">"+data.total_comp+"</div></div><div class=\\"stat\\"><div class=\\"stat-label\\">Verificados</div><div class=\\"stat-value c-green\\">"+data.verificados+"</div></div><div class=\\"stat\\"><div class=\\"stat-label\\">Sin coincidencia</div><div class=\\"stat-value c-red\\">"+(data.total_comp-data.verificados-(data.n_dup||0))+"</div></div><div class=\\"stat\\"><div class=\\"stat-label\\">Repetidos</div><div class=\\"stat-value\\" style=\\"color:#5B21B6\\">"+(data.n_dup||0)+"</div></div>";}\nfunction renderCompResult(comps){document.getElementById("comp-body").innerHTML=comps.map(function(c){var chip=c.duplicado?"<span class=\\"chip chip-d\\" title=\\""+c.duplicado.replace(/"/g,"")+"\\"><span class=\\"dot dot-d\\"></span>Repetido</span>":c.revisar?"<span class=\\"chip chip-a\\" title=\\""+(c.mov_desc||"").replace(/"/g,"")+"\\"><span class=\\"dot dot-a\\"></span>Por monto</span>":c.verificado?"<span class=\\"chip chip-v\\"><span class=\\"dot dot-g\\"></span>Verificado</span>":"<span class=\\"chip chip-r\\"><span class=\\"dot dot-r\\"></span>Sin coincidencia</span>";var monto=c.monto?"$"+c.monto.toLocaleString("es-AR",{minimumFractionDigits:0}):"-";return"<tr class=\\""+(c.duplicado?"dup":c.revisar?"revisar":c.verificado?"verified":"sinmatch")+"\\">"+"<td style=\\"font-size:12px;color:#374151\\">"+c.file+"</td>"+"<td style=\\"color:#6B7280;font-variant-numeric:tabular-nums\\">"+c.fecha+"</td>"+"<td style=\\"text-align:right;font-weight:600;font-variant-numeric:tabular-nums\\">"+monto+"</td>"+"<td style=\\"font-weight:500\\">"+(c.nombre||"-")+"</td>"+"<td>"+chip+"</td>"+"<td><span class=\\"match-tag\\">"+(c.duplicado||c.match||"-")+(c.mov_desc?"<br><span style=\\"color:#6B7280;font-size:10.5px\\">"+c.mov_desc+"</span>":"")+"</span></td>"+"</tr>";}).join("");document.getElementById("results-section").classList.remove("hidden");}\nfunction renderSinMatch(items){var s=document.getElementById("sinmatch-section");if(!items||items.length===0){s.classList.add("hidden");return;}document.getElementById("sinmatch-body").innerHTML=items.map(function(c){var monto=c.monto?"$"+c.monto.toLocaleString("es-AR",{minimumFractionDigits:0}):"-";return"<tr class=\\"sinmatch\\"><td style=\\"font-size:12px\\">"+c.file+"</td><td style=\\"color:#6B7280\\">"+c.fecha+"</td><td style=\\"text-align:right;font-weight:600\\">"+monto+"</td><td>"+(c.nombre||"-")+"</td><td>"+(c.banco_origen||"-")+"</td><td style=\\"font-family:monospace;font-size:11px\\">"+(c.cuit||"-")+"</td></tr>";}).join("");s.classList.remove("hidden");}\nfunction exportFile(fmt){if(!lastCompResult.length)return;var fd=new FormData();fd.append("rows",JSON.stringify(lastRows));fd.append("comp_result",JSON.stringify(lastCompResult));fd.append("sin_match",JSON.stringify(lastSinMatch));fd.append("format",fmt);fd.append("bank_label",lastBankLabel);fd.append("conciliacion_id",lastCid||"");fetch("/exportar",{method:"POST",body:fd}).then(function(r){return r.blob();}).then(function(blob){var url=URL.createObjectURL(blob);var a=document.createElement("a");a.href=url;a.download=fmt==="xlsx"?"conciliacion_"+(lastCid||"borrador")+".xlsx":"conciliacion_"+(lastCid||"borrador")+".pdf";a.click();URL.revokeObjectURL(url);});}\n\n// NUEVA CONCILIACION\nfunction nuevaConciliacion(){\n  xlsxFile=null;compFiles=[];lastCompResult=[];lastRows=[];lastSinMatch=[];lastCid=null;lastBankKey="";lastMoneda="";lastDesc="";conciliacionCerrada=false;\n  renderXlsx();renderComp();check();\n  document.getElementById("inp-desc").value="";\n  document.getElementById("results-section").classList.add("hidden");\n  document.getElementById("sinmatch-section").classList.add("hidden");\n  document.getElementById("dup-section").classList.add("hidden");\n  document.getElementById("error-area").innerHTML="";\n  document.getElementById("prog-wrap").classList.add("hidden");\n  document.getElementById("status-msg").textContent="";\n  var cb=document.getElementById("cerrar-btn");cb.disabled=true;cb.textContent="\\u2714 Cerrar conciliaci\\u00f3n";cb.style.background="";\n  document.getElementById("inp-xlsx").value="";\n  document.getElementById("inp-comp").value="";\n}\n\n// CERRAR CONCILIACION\nasync function cerrarConciliacion(){\n  if(conciliacionCerrada){alert("Esta conciliaci\\u00f3n ya fue cerrada (N\\u00b0 "+lastCid+").");return;}\n  if(!lastCompResult.length){return;}\n  var btn=document.getElementById("cerrar-btn");\n  btn.disabled=true;btn.textContent="Guardando...";\n  try{\n    var resp=await fetch("/cerrar",{method:"POST",headers:{"Content-Type":"application/json"},\n      body:JSON.stringify({bank_key:lastBankKey,comp_result:lastCompResult,verificados:lastCompResult.filter(function(c){return c.verificado;}).length,moneda:lastMoneda,descripcion:lastDesc,rows:lastRows})});\n    var data=await resp.json();\n    if(data.error){alert("Error al cerrar: "+data.error);btn.disabled=false;btn.textContent="\\u2714 Cerrar conciliaci\\u00f3n";return;}\n    lastCid=data.conciliacion_id;\n    conciliacionCerrada=true;\n    btn.textContent="\\u2714 Cerrada \\u2014 N\\u00b0 "+lastCid;\n    btn.style.background="#15803D";\n    document.getElementById("status-msg").textContent="\\u2714 Conciliaci\\u00f3n N\\u00b0 "+lastCid+" guardada en historial.";\n    renderStats({n_dup:lastNDup,conciliacion_id:lastCid,total_comp:lastCompResult.length,verificados:lastCompResult.filter(function(c){return c.verificado;}).length});\n  }catch(e){alert("Error: "+e.message);btn.disabled=false;btn.textContent="\\u2714 Cerrar conciliaci\\u00f3n";}\n}\n\n// REPETIDOS\nfunction renderDup(items){var s=document.getElementById("dup-section");if(!items.length){s.classList.add("hidden");return;}\n  document.getElementById("dup-body").innerHTML=items.map(function(c){var monto=c.monto?"$"+c.monto.toLocaleString("es-AR",{minimumFractionDigits:0}):"-";\n    return \'<tr class="dup"><td style="font-size:12px">\'+c.file+\'</td><td style="color:#6B7280">\'+(c.fecha||"-")+\'</td><td style="text-align:right;font-weight:600">\'+monto+\'</td><td>\'+(c.nombre||"-")+\'</td><td style="color:#5B21B6;font-size:12px">\'+c.duplicado+\'</td></tr>\';}).join("");\n  s.classList.remove("hidden");}\n\n// TRANSFERENCIAS VERIFICADAS\nfunction toggleTransf(){var s=document.getElementById("tv-section");\n  if(s.classList.contains("hidden")){s.classList.remove("hidden");document.getElementById("tv-btn-txt").textContent="Ocultar transferencias";buscarTransf();}\n  else{s.classList.add("hidden");document.getElementById("tv-btn-txt").textContent="Transferencias verificadas";}}\nfunction buscarTransfDebounced(){clearTimeout(tvTimer);tvTimer=setTimeout(buscarTransf,300);}\nasync function buscarTransf(){\n  var q=document.getElementById("tv-q").value||"";\n  var r=await fetch("/transferencias?q="+encodeURIComponent(q));var d=await r.json();\n  document.getElementById("tv-total").textContent="("+d.total+" en la base"+(q?", "+d.rows.length+" coinciden":"")+")";\n  document.getElementById("tv-body").innerHTML=d.rows.map(function(t){\n    var monto=t.monto?"$"+Number(t.monto).toLocaleString("es-AR",{minimumFractionDigits:0}):"-";\n    return \'<tr><td style="color:#6B7280;white-space:nowrap">\'+(t.fecha||"-")+\'</td><td style="text-align:right;font-weight:600">\'+monto+\'</td><td>\'+(t.nombre||"-")+\'</td><td style="font-family:monospace;font-size:11px">\'+(t.cuit||"-")+\'</td><td style="font-size:12px">\'+(t.archivo||"")+\'</td><td style="font-weight:700;color:var(--blue)">N\\u00b0 \'+t.conciliacion_id+\'</td><td style="font-size:11px;color:#6B7280">\'+(t.mov_desc||"")+\'</td></tr>\';\n  }).join("")||\'<tr><td colspan="7" style="text-align:center;color:#9CA3AF;padding:1.5rem">Sin resultados</td></tr>\';\n}\nasync function eliminarConciliacion(cid){\n  if(!confirm("\\u00bfEliminar la conciliaci\\u00f3n N\\u00b0 "+cid+"? Sus transferencias dejan de contar como verificadas."))return;\n  await fetch("/historial/"+cid,{method:"DELETE"});\n  var s=document.getElementById("hist-section");s.classList.add("hidden");cargarHistorial();\n}\n\n// HISTORIAL\nasync function cargarHistorial(){\n  var sec=document.getElementById("hist-section");\n  if(sec.classList.contains("hidden")){sec.classList.remove("hidden");document.getElementById("hist-btn-txt").textContent="Ocultar historial";}\n  else{sec.classList.add("hidden");document.getElementById("hist-btn-txt").textContent="Ver historial";return;}\n  var resp=await fetch("/historial");var data=await resp.json();\n  var fmtMonto=function(m,moneda){if(!m)return"-";var s="$"+parseFloat(m).toLocaleString("es-AR",{minimumFractionDigits:0});return moneda==="PYG"?s+" PYG":s;};\n  var html=data.map(function(r){\n    var badge=r.moneda==="PYG"?"<span style=\\"background:#FEF3C7;color:#92400E;font-size:10px;font-weight:700;padding:2px 7px;border-radius:4px\\">PYG</span>":"<span style=\\"background:#DCFCE7;color:#15803D;font-size:10px;font-weight:700;padding:2px 7px;border-radius:4px\\">ARS</span>";\n    var desc=r.descripcion||"";\n    return"<tr>"+\n      "<td style=\\"font-weight:700;color:var(--blue)\\">N\\u00b0 "+r.id+"</td>"+\n      "<td style=\\"color:#6B7280;font-size:12px\\">"+r.created_at.replace("T"," ").substring(0,16)+"</td>"+\n      "<td>"+r.banco+" "+badge+"</td>"+\n      "<td style=\\"text-align:center\\">"+r.total_comp+"</td>"+\n      "<td style=\\"color:var(--green);font-weight:600;text-align:center\\">"+r.verificados+"</td>"+\n      "<td style=\\"color:var(--red);text-align:center\\">"+r.sin_match+"</td>"+\n      "<td style=\\"font-weight:600\\">"+fmtMonto(r.total_importe,r.moneda)+"</td>"+\n      "<td><div style=\\"display:flex;gap:6px;align-items:center\\">"+\n        "<input type=\\"text\\" value=\\""+desc.replace(/"/g,"&quot;")+"\\" placeholder=\\"Descripci\\u00f3n...\\" style=\\"flex:1;min-width:120px;padding:4px 8px;border:1.5px solid #E4E7ED;border-radius:6px;font-size:12px;font-family:inherit\\" onchange=\\"guardarDesc("+r.id+",this.value)\\">"+\n        "<button onclick=\\"descargarPDFHistorial("+r.id+")\\" style=\\"padding:4px 10px;font-size:12px;border:1.5px solid #B91C1C;color:#B91C1C;background:#fff;border-radius:6px;cursor:pointer;white-space:nowrap;font-family:inherit;font-weight:600\\">PDF</button>"+\n        "<button class=\\"btn-del\\" title=\\"Eliminar conciliaci\\u00f3n\\" onclick=\\"eliminarConciliacion("+r.id+")\\">&#x1F5D1;</button>"+\n      "</div></td>"+\n    "</tr>";\n  }).join("");\n  document.getElementById("hist-body").innerHTML=html||"<tr><td colspan=\\"8\\" style=\\"text-align:center;color:#9CA3AF;padding:2rem\\">Sin registros</td></tr>";\n}\nasync function guardarDesc(cid,desc){\n  await fetch("/historial/"+cid+"/descripcion",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({descripcion:desc})});\n}\nfunction descargarPDFHistorial(cid){\n  var a=document.createElement("a");a.href="/historial/"+cid+"/pdf";a.download="conciliacion_"+cid+".pdf";a.click();\n}\n</script>\n</body>\n</html>'

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
        result=conciliar(plan_path,comp_paths,bank_key)
        moneda = BANK_CONFIG.get(bank_key,{}).get("moneda","ARS")
        result["bank_key"]=bank_key
        result["moneda"]=moneda
        result["descripcion"]=request.form.get("descripcion","")
        log.info("Conciliar %s: %d comprobantes, %d verificados, %d repetidos, %d sin match",
                 bank_key,result["total_comp"],result["verificados"],len(result["duplicados"]),len(result["sin_match"]))
        return jsonify(result)
    except Exception as e:
        log.exception("Error en /conciliar")
        return jsonify({"error":f"{type(e).__name__}: {e}"})

@app.route("/cerrar",methods=["POST"])
def cerrar_route():
    try:
        data=request.json
        bank_key=data.get("bank_key","Ciudad")
        comp_result=data.get("comp_result",[])
        verificados=data.get("verificados",0)
        moneda=data.get("moneda","ARS")
        descripcion=data.get("descripcion","")
        rows=data.get("rows",[])
        prev=cargar_verificadas(); choques=[]
        for c in comp_result:
            if not c.get("verificado"): continue
            p=next((p for p in prev if es_misma_transferencia(c,p)),None)
            if p: choques.append(f"{c['file']} ya verificada en N° {p['conciliacion_id']}")
        if choques: return jsonify({"error":"No se cerro: hay transferencias ya verificadas antes. "+"; ".join(choques[:10])})
        total_importe=sum(c.get("monto") or 0 for c in comp_result if c.get("verificado"))
        comp_data=[{"file":c["file"],"hash":c.get("hash",""),"monto":c.get("monto"),"fecha":c.get("fecha"),"nombre":c.get("nombre")} for c in comp_result]
        cid=guardar_conciliacion(bank_key,comp_data,verificados,moneda,total_importe,comp_result,descripcion)
        log.info("Conciliacion N %s cerrada", cid)
        return jsonify({"ok":True,"conciliacion_id":cid})
    except Exception as e:
        log.exception("Error en /cerrar")
        return jsonify({"error":f"{type(e).__name__}: {e}"})

@app.route("/transferencias",methods=["GET"])
def transferencias_route():
    q=(request.args.get("q") or "").strip()
    conn=get_db()
    sql="SELECT id,conciliacion_id,archivo,monto,fecha,cuit,nombre,banco_origen,referencia,mov_desc,criterio FROM transferencias_verificadas"
    args=[]
    if q:
        num=normalize_monto(q) if re.fullmatch(r"[\d\.,\s\$]+",q) else 0
        sql+=" WHERE nombre LIKE ? OR cuit LIKE ? OR archivo LIKE ? OR referencia LIKE ? OR mov_desc LIKE ?"
        like=f"%{q}%"; args=[like,f"%{clean_cuit(q) or q}%",like,f"%{q.upper()}%",like]
        if num: sql+=" OR ABS(monto-?)<0.5"; args.append(num)
    sql+=" ORDER BY fecha DESC, id DESC LIMIT 300"
    rows=[dict(r) for r in conn.execute(sql,args).fetchall()]
    total=conn.execute("SELECT COUNT(*) FROM transferencias_verificadas").fetchone()[0]
    conn.close()
    return jsonify({"total":total,"rows":rows})

@app.route("/historial/<int:cid>",methods=["DELETE"])
def eliminar_conciliacion(cid):
    conn=get_db()
    conn.execute("DELETE FROM transferencias_verificadas WHERE conciliacion_id=?",(cid,))
    conn.execute("DELETE FROM comp_usados WHERE conciliacion_id=?",(cid,))
    conn.execute("DELETE FROM conciliaciones WHERE id=?",(cid,))
    conn.commit(); conn.close()
    return jsonify({"ok":True})

@app.route("/historial",methods=["GET"])
def historial_route():
    conn=get_db()
    rows=conn.execute("SELECT id,banco,moneda,total_comp,verificados,sin_match,total_importe,descripcion,created_at FROM conciliaciones ORDER BY created_at DESC LIMIT 100").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/historial/<int:cid>/descripcion",methods=["POST"])
def actualizar_descripcion(cid):
    desc=request.json.get("descripcion","")
    conn=get_db()
    conn.execute("UPDATE conciliaciones SET descripcion=? WHERE id=?",(desc,cid))
    conn.commit(); conn.close()
    return jsonify({"ok":True})

@app.route("/historial/<int:cid>/pdf",methods=["GET"])
def historial_pdf(cid):
    conn=get_db()
    row=conn.execute("SELECT banco,moneda,detalle,descripcion FROM conciliaciones WHERE id=?",(cid,)).fetchone()
    conn.close()
    if not row or not row["detalle"]: return "No hay detalle guardado",404
    comp_result=json.loads(row["detalle"])
    bank_label=BANK_CONFIG.get(row["banco"],{}).get("label",row["banco"])
    buf=export_pdf(comp_result,cid,bank_label,descripcion=row["descripcion"] or "")
    return send_file(buf,download_name=f"conciliacion_{cid}.pdf",as_attachment=True,mimetype="application/pdf")

@app.route("/historial/<int:cid>",methods=["GET"])
def historial_detalle(cid):
    conn=get_db()
    row=conn.execute("SELECT detalle FROM conciliaciones WHERE id=?",(cid,)).fetchone()
    conn.close()
    if not row or not row["detalle"]: return jsonify([])
    return jsonify(json.loads(row["detalle"]))

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
