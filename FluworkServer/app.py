# app.py
from flask import Flask, request, jsonify, send_from_directory, Response
from pathlib import Path
import os, json, re, unicodedata
from datetime import datetime, timedelta  # Agregar timedelta
from text_search import TextIndexer, TextSearcher      # NUEVO
from audit_log import registrar_log       # NUEVO (reemplaza la funciÃ³n local)
# --- Lectura de Fluwork ---
from parser import iter_registros, buscar_por_ficha
from serial_index import SerialIndex  # <- NUEVO
# ===================== CONFIG =====================
VERSION = "1.5.0 — 19/03/2026"
# Carpeta donde estÃ¡n las fichas individuales (archivos con nombre numÃ©rico)
DEFAULT_SERVER_ROOT = Path(__file__).resolve().parent
SERVER_ROOT = Path(os.environ.get("FLUWORK_SERVER_ROOT", str(DEFAULT_SERVER_ROOT)))
FICHAS_ROOT = Path(os.environ.get("FLUWORK_FICHAS_ROOT", r"\\MAGENTAOLIVOS\magenta"))
# DÃ³nde guardar la base del Ã­ndice de series
SERIAL_DB = Path(os.environ.get("FLUWORK_SERIAL_DB", str(SERVER_ROOT / "data" / "serial_index.db")))
STATIC_DIR = os.environ.get("FLUWORK_STATIC_DIR", str(SERVER_ROOT / "static"))
TEXT_INDEX_DB = Path(os.environ.get("FLUWORK_TEXT_INDEX_DB", str(SERVER_ROOT / "data" / "fluwork_index.db")))
serial_idx  = SerialIndex(SERIAL_DB, FICHAS_ROOT)  # <- NUEVO
# Ãndice de bÃºsqueda textual y series
TEXT_INDEX_DB.parent.mkdir(parents=True, exist_ok=True)

# Instanciar buscador (read-only)
text_searcher = TextSearcher(TEXT_INDEX_DB)

LOG_PATH = Path("log_consultas.xlsx")
MAX_PAGE = 1000
MAX_VB6_RESULTS = 50000

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path="/static")
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0   # sin cachÃ© para desarrollo

print(">>> STATIC_FOLDER =", os.path.abspath(app.static_folder))
print(">>> FICHAS_ROOT =", str(FICHAS_ROOT))

# ===================== UI / ESTÃTICOS =====================
@app.route("/")
def home():
    # Sirve el index.html directamente desde STATIC_DIR
    return send_from_directory(app.static_folder, "index.html")

@app.route("/favicon.ico")
def favicon():
    try:
        return send_from_directory(app.static_folder, "favicon.ico")
    except Exception:
        return ("", 204)

@app.after_request
def add_no_cache(resp):
    # Evitar que el navegador â€œpegueâ€ versiones viejas durante el dev
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

# ===================== UTILIDADES =====================
def normalizar_nombre(nombre: str) -> str:
    """Quita acentos y pasa a minÃºsculas."""
    return ''.join(
        c for c in unicodedata.normalize('NFD', nombre)
        if unicodedata.category(c) != 'Mn'
    ).lower().strip()

def normalizar_fecha_txt(txt: str) -> str | None:
    m = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\s*$", txt or "")
    if not m:
        return None
    d, mm, y = m.groups()
    if len(y) == 2:
        y = "20" + y if int(y) < 50 else "19" + y
    return f"{int(d):02d}/{int(mm):02d}/{int(y):04d}"

def parse_fecha(txt: str):
    norm = normalizar_fecha_txt(txt)
    if not norm:
        return None
    try:
        return datetime.strptime(norm, "%d/%m/%Y").date()
    except ValueError:
        return None

def fecha_ok(txt, f_desde, f_hasta):
    d = parse_fecha(txt)
    if not d:
        return False
    if f_desde and d < f_desde:
        return False
    if f_hasta and d > f_hasta:
        return False
    return True

def fecha_a_key(reg):
    try:
        return parse_fecha(reg.get("fecha")) or datetime(1900, 1, 1).date()
    except Exception:
        return datetime(1900, 1, 1).date()

# ===================== ENDPOINTS API =====================
# 1) Ficha puntual
@app.get("/api/ficha/<numero>")
def api_ficha(numero):
    reg = buscar_por_ficha(numero)
    registrar_log("anon", "/api/ficha", {"ficha": numero}, 1 if reg else 0)
    if not reg:
        return jsonify({"error": "ficha no encontrada"}), 404
    return jsonify(reg)

# 2) BÃºsqueda flexible (con paginaciÃ³n)
@app.post("/api/buscar")
def api_buscar():
    data = request.get_json(force=True) or {}
    eq   = {k.lower(): v for k, v in data.get("equals", {}).items()}
    con  = {k.lower(): v for k, v in data.get("contains", {}).items()}
    # --- NUEVO: filtro por NÂ° de serie vÃ­a Ã­ndice SQLite ---
    serial_eq  = eq.pop("serial", None)     # si viene en equals â†’ exacto
    serial_con = con.pop("serial", None)    # si viene en contains â†’ prefijo

    allowed_fichas = None
    if serial_eq or serial_con:
        # ActualizaciÃ³n incremental rÃ¡pida del Ã­ndice antes de buscar
        try:
            serial_idx.update_incremental(limit=2000)
        except Exception as e:
            print("WARN serial_idx update:", e)

        if serial_eq:
            allowed_fichas = set(serial_idx.search(serial_eq, exact=True))
        else:
            # por prefijo (rÃ¡pido). Si querÃ©s mÃ¡s laxo, startswith_only=False
            allowed_fichas = set(
                serial_idx.search(serial_con, exact=False, startswith_only=True)
            )

        # Si no hay coincidencias por serie, devolvemos vacÃ­o sin leer index.dat
        if not allowed_fichas:
            return jsonify({"total": 0, "limit": limit, "offset": offset, "items": []})

    f_desde = (datetime.strptime(data["fecha_desde"], "%Y-%m-%d").date()
               if data.get("fecha_desde") else None)
    f_hasta = (datetime.strptime(data["fecha_hasta"], "%Y-%m-%d").date()
               if data.get("fecha_hasta") else None)

    limit  = min(int(data.get("limit", 100)), MAX_PAGE)
    offset = int(data.get("offset", 0))

    resultados = []
    for reg in iter_registros():
        # --- NUEVO: si vino filtro de serie, limitar a esas fichas
        if allowed_fichas is not None and reg.get("ficha") not in allowed_fichas:
            continue
        # filtros equals
        if any(str(reg.get(k, "")).upper() != str(v).upper() for k, v in eq.items()):
            continue
        # filtros contains
        if any(str(v).upper() not in str(reg.get(k, "")).upper() for k, v in con.items()):
            continue
        # fecha dentro de rango (si corresponde)
        if (f_desde or f_hasta) and not fecha_ok(reg.get("fecha", ""), f_desde, f_hasta):
            continue
        resultados.append(reg)

    resultados.sort(key=fecha_a_key, reverse=True)
    total = len(resultados)

    registrar_log("anonimo", "/api/buscar", data, total)

    return jsonify({
        "total": total,
        "limit": limit,
        "offset": offset,
        "items": resultados[offset:offset + limit]
    })

# 3) Autocompletado de tÃ©cnico
@app.get("/api/sugerir_tecnico")
def sugerir_tecnico():
    q = request.args.get("q", "").strip().lower()
    q_norm = normalizar_nombre(q)
    sugerencias = set()
    if len(q) >= 2:
        for reg in iter_registros():
            tecnicos_crudos = reg.get("tecnico", "")
            # Divide por /, \, coma, punto y coma, barra, pipe, espacio
            nombres = re.split(r"[\/,;|]+|\s+", tecnicos_crudos.strip())
            for nombre in nombres:
                nombre_limpio = nombre.strip()
                if not nombre_limpio:
                    continue
                nombre_norm = normalizar_nombre(nombre_limpio)
                if q_norm in nombre_norm and len(nombre_norm) > 1:
                    sugerencias.add(nombre_limpio)
    return jsonify(sorted(sugerencias, key=lambda x: normalizar_nombre(x))[:10])

# Salud del servicio (Ãºtil para pruebas)
@app.get("/health")
def health():
    return jsonify({"ok": True, "static": os.path.abspath(app.static_folder)})

@app.get("/api/version")
def api_version():
    return jsonify({"version": VERSION})
# â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
# NUEVO ENDPOINT (agregar ANTES de app.run)
# â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

@app.post("/api/buscar_avanzada")
def api_buscar_avanzada():
    """
    BÃºsqueda avanzada: texto + serie + filtros estructurados
    Usa Ã­ndice SQLite FTS5 (mucho mÃ¡s rÃ¡pido que index.dat)
    """
    data = request.get_json(force=True) or {}

    # ParÃ¡metros de bÃºsqueda
    # ParÃ¡metros de bÃºsqueda
    problema = data.get("problema", "").strip()
    trabajos = data.get("trabajos", "").strip()
    serie = data.get("serie", "").strip()
    tecnico = data.get("tecnico", "").strip()
    estado = data.get("estado", "").strip()
    telefono = data.get("telefono", "").strip()
    cliente = data.get("cliente", "").strip()
    modelo = data.get("modelo", "").strip()
    email = data.get("email", "").strip()
    confirmacion = data.get("confirmacion", "").strip()
    fecha_desde = data.get("fecha_desde")
    fecha_hasta = data.get("fecha_hasta")
    max_total = int(data.get("max_total", 0))  # 0 = sin tope ("Todas")

    # ¿El usuario puso fecha explícitamente?
    usuario_filtro_fecha = bool(fecha_desde or fecha_hasta)

    # ¿Hay algún filtro activo?
    hay_filtros = any([problema, trabajos, serie, tecnico, estado,
                       telefono, cliente, modelo, email, confirmacion])

    # DEFAULT 60 días: solo cuando NO hay filtros y NO pidió "Todas"
    # Si hay cualquier filtro activo, no limitar por fecha automáticamente
    if not fecha_desde and not fecha_hasta and not hay_filtros and max_total != 0:
        hoy = datetime.now().date()
        hace_60_dias = hoy - timedelta(days=60)
        fecha_desde = hace_60_dias.strftime("%d/%m/%Y")
        print(f"INFO: Aplicando default fecha_desde={fecha_desde}")

    # Límites
    limit = min(int(data.get("limit", 100)), 1000)
    offset = int(data.get("offset", 0))

    # Búsqueda combinada
    try:
        resultados = text_searcher.search_combined(
            problema=problema,
            trabajos=trabajos,
            serie=serie,
            tecnico=tecnico,
            estado=estado,
            telefono=telefono,
            cliente=cliente,
            modelo=modelo,
            email=email,
            confirmacion=confirmacion,
            fecha_desde=fecha_desde,
            fecha_hasta=fecha_hasta,
            limit=limit,
            offset=offset,
            sort_by_date=usuario_filtro_fecha
        )

        # Aplicar tope de resultados (max_total)
        if max_total > 0 and resultados["total"] > max_total:
            resultados["total"] = max_total
            # Si el offset supera el tope, devolver vacío
            if offset >= max_total:
                resultados["items"] = []
            # Si estamos en la última página, truncar items
            elif offset + len(resultados["items"]) > max_total:
                keep = max_total - offset
                resultados["items"] = resultados["items"][:keep]

        registrar_log("anon", "/api/buscar_avanzada", data, resultados["total"])

        return jsonify(resultados)

    except Exception as e:
        print(f"ERROR en bÃºsqueda avanzada: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e), "total": 0, "items": []}), 500

def _vb6_field(value):
    return str(value or "").replace("|", " ").replace("\r", " ").replace("\n", " ").strip()

@app.route("/api/buscar_vb6", methods=["GET", "POST"])
def api_buscar_vb6():
    """
    Endpoint liviano para Fluwork VB6.
    Devuelve texto plano:
      TOTAL|cantidad
      ficha|cliente|telefono|modelo|fecha|estado|tecnico|email|confirmacion
    """
    data = request.get_json(silent=True) or request.form.to_dict() or request.args.to_dict()

    try:
        limit = min(int(data.get("limit", 100)), MAX_VB6_RESULTS)
        offset = int(data.get("offset", 0))

        resultados = text_searcher.search_combined(
            problema=data.get("problema", "").strip(),
            trabajos=data.get("trabajos", "").strip(),
            serie=data.get("serie", "").strip(),
            tecnico=data.get("tecnico", "").strip(),
            estado=data.get("estado", "").strip(),
            telefono=data.get("telefono", "").strip(),
            cliente=data.get("cliente", "").strip(),
            modelo=data.get("modelo", "").strip(),
            email=data.get("email", "").strip(),
            confirmacion=data.get("confirmacion", "").strip(),
            fecha_desde=data.get("fecha_desde") or None,
            fecha_hasta=data.get("fecha_hasta") or None,
            limit=limit,
            offset=offset,
            sort_by_date=bool(data.get("fecha_desde") or data.get("fecha_hasta"))
        )

        lines = [f"TOTAL|{resultados['total']}"]
        for item in resultados["items"]:
            lines.append("|".join([
                _vb6_field(item.get("ficha")),
                _vb6_field(item.get("fullname")),
                _vb6_field(item.get("telefono")),
                _vb6_field(item.get("modelo")),
                _vb6_field(item.get("fecha")),
                _vb6_field(item.get("estado")),
                _vb6_field(item.get("tecnico")),
                _vb6_field(item.get("email")),
                _vb6_field(item.get("confirmacion")),
            ]))

        registrar_log("vb6", "/api/buscar_vb6", data, resultados["total"])
        body = "\r\n".join(lines) + "\r\n"
        return Response(body.encode("cp1252", errors="replace"), mimetype="text/plain; charset=windows-1252")

    except Exception as e:
        body = "ERROR|" + _vb6_field(e) + "\r\n"
        return Response(body.encode("cp1252", errors="replace"), status=500, mimetype="text/plain; charset=windows-1252")

@app.route("/api/reindex_ficha_vb6", methods=["GET", "POST"])
def api_reindex_ficha_vb6():
    data = request.get_json(silent=True) or request.form.to_dict() or request.args.to_dict()
    ficha = re.sub(r"\D", "", str(data.get("ficha", "")))

    if not ficha:
        body = "ERROR|ficha invalida\r\n"
        return Response(body.encode("cp1252", errors="replace"), status=400, mimetype="text/plain; charset=windows-1252")

    try:
        indexer = TextIndexer(TEXT_INDEX_DB, FICHAS_ROOT)
        updated = indexer.update_one(ficha)
        if updated <= 0:
            body = f"ERROR|ficha no encontrada|{_vb6_field(ficha)}\r\n"
            return Response(body.encode("cp1252", errors="replace"), status=404, mimetype="text/plain; charset=windows-1252")

        body = f"OK|{_vb6_field(ficha)}\r\n"
        return Response(body.encode("cp1252", errors="replace"), mimetype="text/plain; charset=windows-1252")

    except Exception as e:
        body = "ERROR|" + _vb6_field(e) + "\r\n"
        return Response(body.encode("cp1252", errors="replace"), status=500, mimetype="text/plain; charset=windows-1252")
# Endpoint de diagnÃ³stico (opcional)
@app.get("/api/text_index/status")
def text_index_status():
    """InformaciÃ³n sobre el estado del Ã­ndice"""
    try:
        stats = text_searcher.get_stats()
        return jsonify({
            "ok": True,
            "db_exists": TEXT_INDEX_DB.exists(),
            "db_path": str(TEXT_INDEX_DB),
            **stats
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})
# ===================== RUN =====================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False, use_reloader=False)
