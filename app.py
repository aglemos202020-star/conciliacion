"""Caja de boletería — backend Flask + SQLite para Railway."""
import os, sqlite3, secrets, threading, time
from datetime import datetime
from functools import wraps
from zoneinfo import ZoneInfo
from flask import Flask, g, jsonify, request, session, send_from_directory
from werkzeug.security import generate_password_hash, check_password_hash

DB_PATH = os.environ.get("DB_PATH", os.path.join(os.path.dirname(__file__), "caja.db"))
TZ = ZoneInfo(os.environ.get("APP_TZ", "America/Argentina/Buenos_Aires"))
MASTER_USER = os.environ.get("MASTER_USER", "lemos").strip().lower()
MASTER_PASS = os.environ.get("MASTER_PASS", "")

app = Flask(__name__, static_folder="static")
app.config.update(
    SECRET_KEY=os.environ.get("SECRET_KEY") or secrets.token_hex(32),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE", "1") == "1",
    PERMANENT_SESSION_LIFETIME=60 * 60 * 14,
)

GRUPOS = ("efectivo", "tarjeta", "transferencia", "otro")
TIPOS = ("venta", "devolucion", "ingreso", "egreso")

# ---------------- base de datos ----------------
ESQUEMA = """
CREATE TABLE IF NOT EXISTS agencias (id INTEGER PRIMARY KEY, codigo TEXT UNIQUE NOT NULL, activa INTEGER NOT NULL DEFAULT 1);
CREATE TABLE IF NOT EXISTS usuarios (id INTEGER PRIMARY KEY, usuario TEXT UNIQUE NOT NULL COLLATE NOCASE, nombre TEXT NOT NULL,
  pass_hash TEXT NOT NULL, rol TEXT NOT NULL CHECK (rol IN ('master','vendedor')), agencia_id INTEGER REFERENCES agencias(id),
  activo INTEGER NOT NULL DEFAULT 1);
CREATE TABLE IF NOT EXISTS empresas (id INTEGER PRIMARY KEY, nombre TEXT UNIQUE NOT NULL COLLATE NOCASE, activa INTEGER NOT NULL DEFAULT 1);
CREATE TABLE IF NOT EXISTS medios (id INTEGER PRIMARY KEY, nombre TEXT UNIQUE NOT NULL COLLATE NOCASE, grupo TEXT NOT NULL,
  efectivo INTEGER NOT NULL DEFAULT 0, activo INTEGER NOT NULL DEFAULT 1);
CREATE TABLE IF NOT EXISTS cajas (id INTEGER PRIMARY KEY, agencia_id INTEGER NOT NULL REFERENCES agencias(id), numero INTEGER NOT NULL,
  estado TEXT NOT NULL, abierta_en TEXT NOT NULL, abierta_por INTEGER, fondo INTEGER NOT NULL DEFAULT 0,
  cerrada_en TEXT, cerrada_por INTEGER, esperado INTEGER, contado INTEGER, dif INTEGER, fondo_sig INTEGER, obs TEXT,
  tot_ventas INTEGER, tot_devol INTEGER, UNIQUE (agencia_id, numero));
CREATE UNIQUE INDEX IF NOT EXISTS una_abierta ON cajas(agencia_id) WHERE estado = 'abierta';
CREATE TABLE IF NOT EXISTS ops (id INTEGER PRIMARY KEY, caja_id INTEGER NOT NULL REFERENCES cajas(id), tipo TEXT NOT NULL,
  empresa_id INTEGER, medio_id INTEGER, concepto TEXT, boleto TEXT, cupon TEXT, importe INTEGER NOT NULL,
  usuario_id INTEGER NOT NULL, creado TEXT NOT NULL, anulada INTEGER NOT NULL DEFAULT 0, anulada_por INTEGER, anulada_en TEXT);
CREATE INDEX IF NOT EXISTS ops_caja ON ops(caja_id);
CREATE TABLE IF NOT EXISTS pendientes (id INTEGER PRIMARY KEY, usuario_id INTEGER NOT NULL REFERENCES usuarios(id),
  agencia_id INTEGER NOT NULL REFERENCES agencias(id), empresa_id INTEGER NOT NULL REFERENCES empresas(id), boleto TEXT NOT NULL,
  importe INTEGER NOT NULL, op_egreso_id INTEGER REFERENCES ops(id), creado TEXT NOT NULL,
  estado TEXT NOT NULL DEFAULT 'pendiente' CHECK (estado IN ('pendiente','cobrado','anulado')),
  op_cobro_id INTEGER REFERENCES ops(id), cobrado_en TEXT, anulado_en TEXT);
CREATE INDEX IF NOT EXISTS pend_usuario ON pendientes(usuario_id, estado);
"""

_lock = threading.Lock()  # SQLite: serializamos las escrituras


def conectar():
    con = sqlite3.connect(DB_PATH, timeout=15)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con


def db():
    if "db" not in g:
        g.db = conectar()
    return g.db


@app.teardown_appcontext
def cerrar_db(_):
    con = g.pop("db", None)
    if con:
        con.close()


def ahora():
    return datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")


def iniciar_db():
    d = os.path.dirname(DB_PATH)
    if d:
        os.makedirs(d, exist_ok=True)
    con = conectar()
    con.execute("PRAGMA journal_mode = WAL")
    con.executescript(ESQUEMA)
    cols = {r["name"] for r in con.execute("PRAGMA table_info(ops)")}  # actualizar bases ya existentes
    if "subtipo" not in cols:
        con.execute("ALTER TABLE ops ADD COLUMN subtipo TEXT")
    if "pendiente_id" not in cols:
        con.execute("ALTER TABLE ops ADD COLUMN pendiente_id INTEGER")
    if not con.execute("SELECT 1 FROM agencias").fetchone():
        con.executemany("INSERT INTO agencias (codigo) VALUES (?)", [("219",), ("B22",), ("CCR",)])
    if not con.execute("SELECT 1 FROM empresas").fetchone():
        con.executemany("INSERT INTO empresas (nombre) VALUES (?)",
                        [(n,) for n in ["La Santaniana", "Guaireña", "Yuteña", "Expreso Sur", "JC", "20 de Junio"]])
    if not con.execute("SELECT 1 FROM medios").fetchone():
        con.executemany("INSERT INTO medios (nombre, grupo, efectivo) VALUES (?,?,?)",
                        [("Efectivo", "efectivo", 1), ("Débito", "tarjeta", 0), ("Crédito", "tarjeta", 0),
                         ("Transferencia", "transferencia", 0), ("Mercado Pago", "otro", 0)])
    master = con.execute("SELECT id FROM usuarios WHERE usuario = ?", (MASTER_USER,)).fetchone()
    if not master:
        clave = MASTER_PASS or "lemos123"
        if not MASTER_PASS:
            print("AVISO: falta MASTER_PASS; el master se creó con la contraseña lemos123. Cambiala.", flush=True)
        con.execute("INSERT INTO usuarios (usuario, nombre, pass_hash, rol) VALUES (?,?,?, 'master')",
                    (MASTER_USER, MASTER_USER.capitalize(), generate_password_hash(clave)))
    elif MASTER_PASS:  # la variable de Railway manda: sirve para recuperar el acceso
        con.execute("UPDATE usuarios SET pass_hash = ?, activo = 1, rol = 'master' WHERE id = ?",
                    (generate_password_hash(MASTER_PASS), master["id"]))
    con.commit()
    con.close()


# ---------------- utilidades ----------------
class Error(Exception):
    pass


def centavos(v, cero=False):
    try:
        c = round(float(v) * 100)
    except (TypeError, ValueError):
        raise Error("Importe inválido.")
    if c < 0 or (c == 0 and not cero) or c > 10**13:
        raise Error("Importe inválido.")
    return c


def texto(v, maximo=120):
    return str(v or "").strip()[:maximo]


def usuario_actual():
    uid = session.get("uid")
    if not uid:
        return None
    u = db().execute("SELECT * FROM usuarios WHERE id = ? AND activo = 1", (uid,)).fetchone()
    if not u:
        return None
    if u["rol"] == "vendedor":
        a = db().execute("SELECT activa FROM agencias WHERE id = ?", (u["agencia_id"],)).fetchone()
        if not a or not a["activa"]:
            return None
    return u


def requiere_sesion(f):
    @wraps(f)
    def envoltura(*a, **k):
        if request.method == "POST" and not request.is_json:
            return jsonify(error="Pedido inválido."), 400
        u = usuario_actual()
        if not u:
            session.clear()
            return jsonify(error="sesion"), 401
        g.u = u
        return f(*a, **k)
    return envoltura


def es_master():
    return g.u["rol"] == "master"


def agencia_permitida(ag_id):
    ag_id = int(ag_id or 0)
    if not es_master():
        return g.u["agencia_id"]
    a = db().execute("SELECT id FROM agencias WHERE id = ? AND activa = 1", (ag_id,)).fetchone()
    if not a:
        raise Error("Elegí una agencia activa.")
    return ag_id


def caja_abierta(con, ag_id, uid):
    c = con.execute("SELECT * FROM cajas WHERE agencia_id = ? AND estado = 'abierta'", (ag_id,)).fetchone()
    if c:
        return c
    n = con.execute("SELECT COALESCE(MAX(numero), 0) + 1 FROM cajas WHERE agencia_id = ?", (ag_id,)).fetchone()[0]
    con.execute("INSERT INTO cajas (agencia_id, numero, estado, abierta_en, abierta_por, fondo) VALUES (?,?, 'abierta', ?,?,0)",
                (ag_id, n, ahora(), uid))
    return con.execute("SELECT * FROM cajas WHERE agencia_id = ? AND estado = 'abierta'", (ag_id,)).fetchone()


def totales(con, caja):
    rows = con.execute("""SELECT o.tipo, o.importe, COALESCE(m.efectivo, 0) ef FROM ops o LEFT JOIN medios m ON m.id = o.medio_id
                          WHERE o.caja_id = ? AND o.anulada = 0""", (caja["id"],)).fetchall()
    t = dict(ventas=0, ventas_ef=0, devol=0, ingresos=0, egresos=0)
    for r in rows:
        if r["tipo"] == "venta":
            t["ventas"] += r["importe"]
            if r["ef"]:
                t["ventas_ef"] += r["importe"]
        elif r["tipo"] == "devolucion":
            t["devol"] += r["importe"]
        elif r["tipo"] == "ingreso":
            t["ingresos"] += r["importe"]
        else:
            t["egresos"] += r["importe"]
    t["esperado"] = caja["fondo"] + t["ventas_ef"] - t["devol"] + t["ingresos"] - t["egresos"]
    return t


P = lambda c: None if c is None else c / 100  # centavos → pesos


def op_json(o):
    return dict(id=o["id"], cajaId=o["caja_id"], tipo=o["tipo"], empresaId=o["empresa_id"], medioId=o["medio_id"],
                concepto=o["concepto"], boleto=o["boleto"], cupon=o["cupon"], importe=P(o["importe"]),
                usuarioId=o["usuario_id"], creado=o["creado"], anulada=bool(o["anulada"]),
                anuladaPor=o["anulada_por"], anuladaEn=o["anulada_en"], subtipo=o["subtipo"], pendienteId=o["pendiente_id"])


def pend_json(p):
    return dict(id=p["id"], agenciaId=p["agencia_id"], empresaId=p["empresa_id"], boleto=p["boleto"], importe=P(p["importe"]),
                creado=p["creado"], estado=p["estado"], cobradoEn=p["cobrado_en"])


def pendientes_de_cierre(con, caja, uid):
    """Boletos pendientes del usuario que seguían sin cobrar al momento de cerrar esa caja."""
    return [pend_json(p) for p in con.execute(
        """SELECT * FROM pendientes WHERE usuario_id = ? AND agencia_id = ? AND creado <= ?
           AND (estado = 'pendiente' OR (estado = 'cobrado' AND cobrado_en > ?) OR (estado = 'anulado' AND anulado_en > ?))
           ORDER BY creado""", (uid, caja["agencia_id"], caja["cerrada_en"], caja["cerrada_en"], caja["cerrada_en"]))]


def estado():
    con = db()
    master = es_master()
    with _lock:  # garantizar una caja abierta por cada agencia visible
        ags = [r["id"] for r in con.execute("SELECT id FROM agencias WHERE activa = 1")] if master else [g.u["agencia_id"]]
        for a in ags:
            caja_abierta(con, a, g.u["id"])
        con.commit()
    agencias = [dict(id=r["id"], codigo=r["codigo"], activa=bool(r["activa"]))
                for r in con.execute("SELECT * FROM agencias ORDER BY codigo")]
    usuarios = []
    for r in con.execute("SELECT * FROM usuarios ORDER BY rol, nombre"):
        u = dict(id=r["id"], nombre=r["nombre"], rol=r["rol"], agenciaId=r["agencia_id"], activo=bool(r["activo"]))
        if master or r["id"] == g.u["id"]:
            u["usuario"] = r["usuario"]
        usuarios.append(u)
    empresas = [dict(id=r["id"], nombre=r["nombre"], activa=bool(r["activa"]))
                for r in con.execute("SELECT * FROM empresas ORDER BY nombre")]
    medios = [dict(id=r["id"], nombre=r["nombre"], grupo=r["grupo"], efectivo=bool(r["efectivo"]), activo=bool(r["activo"]))
              for r in con.execute("SELECT * FROM medios ORDER BY id")]
    filtro, args = ("", ()) if master else ("WHERE agencia_id = ?", (g.u["agencia_id"],))
    cajas, abiertas = [], []
    for r in con.execute(f"SELECT * FROM cajas {filtro} ORDER BY id", args):
        cajas.append(dict(id=r["id"], agenciaId=r["agencia_id"], numero=r["numero"], estado=r["estado"],
                          abiertaEn=r["abierta_en"], abiertaPor=r["abierta_por"], fondo=P(r["fondo"]),
                          cerradaEn=r["cerrada_en"], cerradaPor=r["cerrada_por"], esperado=P(r["esperado"]),
                          contado=P(r["contado"]), dif=P(r["dif"]), fondoSig=P(r["fondo_sig"]), obs=r["obs"] or "",
                          totVentas=P(r["tot_ventas"]), totDevol=P(r["tot_devol"])))
        if r["estado"] == "abierta":
            abiertas.append(r["id"])
    ops = []
    if abiertas:
        q = ",".join("?" * len(abiertas))
        ops = [op_json(o) for o in con.execute(f"SELECT * FROM ops WHERE caja_id IN ({q}) ORDER BY id", abiertas)]
    pend = [pend_json(p) for p in con.execute(
        "SELECT * FROM pendientes WHERE usuario_id = ? AND estado = 'pendiente' ORDER BY creado", (g.u["id"],))]
    yo = dict(id=g.u["id"], usuario=g.u["usuario"], nombre=g.u["nombre"], rol=g.u["rol"], agenciaId=g.u["agencia_id"])
    return dict(usuario=yo, estado=dict(agencias=agencias, usuarios=usuarios, empresas=empresas, medios=medios,
                                        cajas=cajas, ops=ops, pendientes=pend))


# ---------------- rutas ----------------
@app.get("/")
def inicio():
    r = send_from_directory(app.static_folder, "app.html")
    r.headers["Cache-Control"] = "no-cache"
    return r


@app.get("/salud")
def salud():
    return "ok"


_intentos = {}


@app.post("/api/login")
def login():
    if not request.is_json:
        return jsonify(error="Pedido inválido."), 400
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "").split(",")[0].strip()
    ahora_s = time.time()
    lista = [t for t in _intentos.get(ip, []) if ahora_s - t < 600]
    if len(lista) >= 10:
        return jsonify(error="Demasiados intentos. Esperá unos minutos."), 429
    d = request.get_json(silent=True) or {}
    u = db().execute("SELECT * FROM usuarios WHERE usuario = ? AND activo = 1", (texto(d.get("usuario"), 60),)).fetchone()
    if not u or not check_password_hash(u["pass_hash"], str(d.get("pass") or "")):
        lista.append(ahora_s)
        _intentos[ip] = lista
        return jsonify(error="Usuario o contraseña incorrectos."), 401
    if u["rol"] == "vendedor":
        a = db().execute("SELECT activa FROM agencias WHERE id = ?", (u["agencia_id"],)).fetchone()
        if not a or not a["activa"]:
            return jsonify(error="Tu agencia está desactivada. Consultá con el master."), 403
    _intentos.pop(ip, None)
    session.clear()
    session.permanent = True
    session["uid"] = u["id"]
    g.u = u
    return jsonify(estado())


@app.post("/api/salir")
def salir():
    session.clear()
    return jsonify(ok=True)


@app.get("/api/estado")
@requiere_sesion
def api_estado():
    return jsonify(estado())


@app.get("/api/caja/<int:caja_id>/ops")
@requiere_sesion
def api_ops_caja(caja_id):
    c = db().execute("SELECT * FROM cajas WHERE id = ?", (caja_id,)).fetchone()
    if not c or (not es_master() and c["agencia_id"] != g.u["agencia_id"]):
        return jsonify(error="Caja no encontrada."), 404
    pend = pendientes_de_cierre(db(), c, g.u["id"]) if c["estado"] == "cerrada" else []
    return jsonify(ops=[op_json(o) for o in db().execute("SELECT * FROM ops WHERE caja_id = ? ORDER BY id", (caja_id,))],
                   pendientes=pend)


@app.post("/api/accion")
@requiere_sesion
def api_accion():
    d = request.get_json(silent=True) or {}
    accion = d.get("accion")
    fn = ACCIONES.get(accion)
    if not fn:
        return jsonify(error="Acción desconocida."), 400
    if accion in SOLO_MASTER and not es_master():
        return jsonify(error="Solo el master puede hacer esto."), 403
    con = db()
    try:
        with _lock:
            extra = fn(con, d) or {}
            con.commit()
    except Error as e:
        con.rollback()
        return jsonify(error=str(e)), 400
    except sqlite3.IntegrityError:
        con.rollback()
        return jsonify(error="Ese dato ya existe."), 400
    g.u = usuario_actual() or g.u
    res = estado()
    res.update(extra)
    return jsonify(res)


# ---------------- acciones ----------------
def a_op(con, d):
    tipo = d.get("tipo")
    if tipo not in TIPOS:
        raise Error("Tipo de operación inválido.")
    ag = agencia_permitida(d.get("agenciaId"))
    if tipo == "egreso" and d.get("subtipo") == "pendiente":
        return alta_pendiente(con, d, ag)
    if tipo == "ingreso" and d.get("subtipo") == "cobro":
        return cobro_pendiente(con, d, ag)
    importe = centavos(d.get("importe"))
    empresa = medio = concepto = boleto = cupon = None
    if tipo in ("venta", "devolucion"):
        empresa = con.execute("SELECT id FROM empresas WHERE id = ? AND activa = 1", (d.get("empresaId"),)).fetchone()
        if not empresa:
            raise Error("Elegí la empresa.")
        empresa = empresa["id"]
        boleto = texto(d.get("boleto"), 30)
        if not boleto:
            raise Error("Cargá el número de boleto.")
    if tipo == "venta":
        m = con.execute("SELECT * FROM medios WHERE id = ? AND activo = 1", (d.get("medioId"),)).fetchone()
        if not m:
            raise Error("Elegí el modo de pago.")
        medio = m["id"]
        if m["grupo"] == "tarjeta":
            cupon = texto(d.get("cupon"), 30)
            if not cupon:
                raise Error("Cargá el número de cupón de la tarjeta.")
    elif tipo == "devolucion":
        medio = con.execute("SELECT id FROM medios WHERE efectivo = 1 ORDER BY id LIMIT 1").fetchone()["id"]
    else:
        concepto = texto(d.get("concepto"), 120)
        if not concepto:
            raise Error("Escribí el concepto del movimiento.")
    c = caja_abierta(con, ag, g.u["id"])
    con.execute("""INSERT INTO ops (caja_id, tipo, empresa_id, medio_id, concepto, boleto, cupon, importe, usuario_id, creado)
                   VALUES (?,?,?,?,?,?,?,?,?,?)""", (c["id"], tipo, empresa, medio, concepto, boleto, cupon, importe, g.u["id"], ahora()))


def alta_pendiente(con, d, ag):
    """Egreso varios: boleto que se entregó sin cobrar. Queda en el registro del usuario hasta que se cobre."""
    e = con.execute("SELECT * FROM empresas WHERE id = ? AND activa = 1", (d.get("empresaId"),)).fetchone()
    if not e:
        raise Error("Elegí la empresa del boleto pendiente.")
    boleto = texto(d.get("boleto"), 30)
    if not boleto:
        raise Error("Cargá el número de boleto pendiente.")
    if con.execute("SELECT 1 FROM pendientes WHERE usuario_id = ? AND empresa_id = ? AND boleto = ? AND estado = 'pendiente'",
                   (g.u["id"], e["id"], boleto)).fetchone():
        raise Error("Ese boleto ya está en tus pendientes.")
    importe = centavos(d.get("importe"))
    c = caja_abierta(con, ag, g.u["id"])
    t = ahora()
    cur = con.execute("""INSERT INTO ops (caja_id, tipo, subtipo, empresa_id, concepto, boleto, importe, usuario_id, creado)
                         VALUES (?, 'egreso', 'pendiente', ?,?,?,?,?,?)""",
                      (c["id"], e["id"], f"Boleto pendiente de pago: {e['nombre']}, boleto {boleto}", boleto, importe, g.u["id"], t))
    op_id = cur.lastrowid
    cur = con.execute("""INSERT INTO pendientes (usuario_id, agencia_id, empresa_id, boleto, importe, op_egreso_id, creado)
                         VALUES (?,?,?,?,?,?,?)""", (g.u["id"], ag, e["id"], boleto, importe, op_id, t))
    con.execute("UPDATE ops SET pendiente_id = ? WHERE id = ?", (cur.lastrowid, op_id))


def cobro_pendiente(con, d, ag):
    """Ingreso varios: se cobra un boleto pendiente y sale del registro."""
    p = con.execute("""SELECT p.*, e.nombre empresa FROM pendientes p JOIN empresas e ON e.id = p.empresa_id
                       WHERE p.id = ? AND p.usuario_id = ? AND p.estado = 'pendiente'""", (d.get("pendienteId"), g.u["id"])).fetchone()
    if not p:
        raise Error("Elegí uno de tus boletos pendientes.")
    c = caja_abierta(con, ag, g.u["id"])
    t = ahora()
    cur = con.execute("""INSERT INTO ops (caja_id, tipo, subtipo, empresa_id, concepto, boleto, importe, usuario_id, creado, pendiente_id)
                         VALUES (?, 'ingreso', 'cobro', ?,?,?,?,?,?,?)""",
                      (c["id"], p["empresa_id"], f"Cobro de boleto pendiente: {p['empresa']}, boleto {p['boleto']}", p["boleto"],
                       p["importe"], g.u["id"], t, p["id"]))
    con.execute("UPDATE pendientes SET estado = 'cobrado', op_cobro_id = ?, cobrado_en = ? WHERE id = ?", (cur.lastrowid, t, p["id"]))


def a_fondo(con, d):
    if not es_master():
        raise Error("Solo el master puede cambiar el fondo.")
    c = caja_abierta(con, agencia_permitida(d.get("agenciaId")), g.u["id"])
    con.execute("UPDATE cajas SET fondo = ? WHERE id = ?", (centavos(d.get("fondo"), True), c["id"]))


def a_cierre(con, d):
    c = caja_abierta(con, agencia_permitida(d.get("agenciaId")), g.u["id"])
    if d.get("cajaId") and int(d["cajaId"]) != c["id"]:
        raise Error("La caja cambió mientras cerrabas. Revisá y volvé a intentar.")
    contado, fondo_sig = centavos(d.get("contado"), True), centavos(d.get("fondoSig") or 0, True)
    if fondo_sig > contado:
        raise Error("El fondo que queda no puede ser mayor que el efectivo contado.")
    t = totales(con, c)
    con.execute("""UPDATE cajas SET estado = 'cerrada', cerrada_en = ?, cerrada_por = ?, esperado = ?, contado = ?, dif = ?,
                   fondo_sig = ?, obs = ?, tot_ventas = ?, tot_devol = ? WHERE id = ?""",
                (ahora(), g.u["id"], t["esperado"], contado, contado - t["esperado"], fondo_sig, texto(d.get("obs"), 500),
                 t["ventas"], t["devol"], c["id"]))
    n = c["numero"] + 1
    con.execute("INSERT INTO cajas (agencia_id, numero, estado, abierta_en, abierta_por, fondo) VALUES (?,?, 'abierta', ?,?,?)",
                (c["agencia_id"], n, ahora(), g.u["id"], fondo_sig))
    ops = [op_json(o) for o in con.execute("SELECT * FROM ops WHERE caja_id = ? ORDER BY id", (c["id"],))]
    cerrada = con.execute("SELECT * FROM cajas WHERE id = ?", (c["id"],)).fetchone()
    return dict(cajaId=c["id"], numeroNuevo=n, opsCaja=ops, pendCaja=pendientes_de_cierre(con, cerrada, g.u["id"]))


def a_anular(con, d):
    o = con.execute("""SELECT o.*, c.estado, c.agencia_id FROM ops o JOIN cajas c ON c.id = o.caja_id WHERE o.id = ?""",
                    (d.get("id"),)).fetchone()
    if not o or o["anulada"]:
        raise Error("La operación no existe o ya está anulada.")
    if o["estado"] != "abierta":
        raise Error("Solo se pueden anular operaciones de la caja abierta.")
    if not es_master() and (o["usuario_id"] != g.u["id"] or o["agencia_id"] != g.u["agencia_id"]):
        raise Error("Solo podés anular operaciones que cargaste vos.")
    if o["subtipo"] == "pendiente" and o["pendiente_id"]:
        p = con.execute("SELECT estado FROM pendientes WHERE id = ?", (o["pendiente_id"],)).fetchone()
        if p and p["estado"] == "cobrado":
            raise Error("Ese boleto pendiente ya se cobró. Primero anulá el cobro.")
        con.execute("UPDATE pendientes SET estado = 'anulado', anulado_en = ? WHERE id = ?", (ahora(), o["pendiente_id"]))
    elif o["subtipo"] == "cobro" and o["pendiente_id"]:  # vuelve a quedar pendiente
        con.execute("UPDATE pendientes SET estado = 'pendiente', op_cobro_id = NULL, cobrado_en = NULL WHERE id = ?",
                    (o["pendiente_id"],))
    con.execute("UPDATE ops SET anulada = 1, anulada_por = ?, anulada_en = ? WHERE id = ?", (g.u["id"], ahora(), o["id"]))


def a_pass(con, d):
    p = str(d.get("pass") or "")
    if len(p) < 4:
        raise Error("La contraseña debe tener al menos 4 caracteres.")
    u = con.execute("SELECT * FROM usuarios WHERE id = ?", (d.get("id"),)).fetchone()
    if not u:
        raise Error("Usuario inexistente.")
    if u["rol"] == "master" and u["usuario"] == MASTER_USER and MASTER_PASS:
        raise Error("La contraseña del master se cambia desde la variable MASTER_PASS en Railway.")
    con.execute("UPDATE usuarios SET pass_hash = ? WHERE id = ?", (generate_password_hash(p), u["id"]))
    return dict(mensaje=f"Contraseña de {u['nombre']} actualizada.")


def a_nuevousr(con, d):
    nombre, usuario, p = texto(d.get("nombre"), 60), texto(d.get("usuario"), 40).lower(), str(d.get("pass") or "")
    if not nombre or not usuario or len(p) < 4:
        raise Error("Completá nombre, usuario y una contraseña de al menos 4 caracteres.")
    if not con.execute("SELECT 1 FROM agencias WHERE id = ?", (d.get("agenciaId"),)).fetchone():
        raise Error("Asigná una agencia.")
    if con.execute("SELECT 1 FROM usuarios WHERE usuario = ?", (usuario,)).fetchone():
        raise Error("Ese nombre de usuario ya existe.")
    con.execute("INSERT INTO usuarios (usuario, nombre, pass_hash, rol, agencia_id) VALUES (?,?,?, 'vendedor', ?)",
                (usuario, nombre, generate_password_hash(p), d.get("agenciaId")))


def a_usrag(con, d):
    if not con.execute("SELECT 1 FROM agencias WHERE id = ?", (d.get("agenciaId"),)).fetchone():
        raise Error("Agencia inexistente.")
    con.execute("UPDATE usuarios SET agencia_id = ? WHERE id = ? AND rol = 'vendedor'", (d.get("agenciaId"), d.get("id")))


def a_togusr(con, d):
    con.execute("UPDATE usuarios SET activo = 1 - activo WHERE id = ? AND rol = 'vendedor'", (d.get("id"),))


def a_nuevaag(con, d):
    codigo = texto(d.get("codigo"), 12).upper()
    if not codigo:
        raise Error("Escribí el código de la agencia.")
    con.execute("INSERT INTO agencias (codigo) VALUES (?)", (codigo,))


def a_togag(con, d):
    con.execute("UPDATE agencias SET activa = 1 - activa WHERE id = ?", (d.get("id"),))


def a_nuevaemp(con, d):
    n = texto(d.get("nombre"), 60)
    if not n:
        raise Error("Escribí el nombre de la empresa.")
    con.execute("INSERT INTO empresas (nombre) VALUES (?)", (n,))


def a_togemp(con, d):
    con.execute("UPDATE empresas SET activa = 1 - activa WHERE id = ?", (d.get("id"),))


def a_nuevomed(con, d):
    n = texto(d.get("nombre"), 40)
    if not n:
        raise Error("Escribí el nombre del modo de pago.")
    grupo = d.get("grupo") if d.get("grupo") in ("tarjeta", "transferencia", "otro") else "otro"
    con.execute("INSERT INTO medios (nombre, grupo) VALUES (?,?)", (n, grupo))


def a_togmed(con, d):
    con.execute("UPDATE medios SET activo = 1 - activo WHERE id = ? AND efectivo = 0", (d.get("id"),))


ACCIONES = dict(op=a_op, fondo=a_fondo, cierre=a_cierre, anular=a_anular, pass_=a_pass, nuevousr=a_nuevousr, usrag=a_usrag,
                togusr=a_togusr, nuevaag=a_nuevaag, togag=a_togag, nuevaemp=a_nuevaemp, togemp=a_togemp,
                nuevomed=a_nuevomed, togmed=a_togmed)
ACCIONES["pass"] = ACCIONES.pop("pass_")
SOLO_MASTER = {"fondo", "pass", "nuevousr", "usrag", "togusr", "nuevaag", "togag", "nuevaemp", "togemp", "nuevomed", "togmed"}

iniciar_db()

if __name__ == "__main__":
    app.config["SESSION_COOKIE_SECURE"] = False
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))
