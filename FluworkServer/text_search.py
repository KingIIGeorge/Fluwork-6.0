# text_search_optimizado.py
"""
Versión OPTIMIZADA - NO carga index.dat en memoria
Para incremental usa metadata ya guardada en DB
Solo para build_full carga el index.dat
"""
from pathlib import Path
import sqlite3
import re
import time
from datetime import datetime
from audit_log import registrar_log
from parser import ENC, FIELDS, WIDTH

ENCODING_FICHA = "cp1252"

def normalize_serial(s: str) -> str:
    s = s.upper().strip()
    s = re.sub(r"^[\-\/]+|[\-\/]+$", "", s)
    s = re.sub(r"[^A-Z0-9\-\/]", "", s)
    return s

def _to_sortable(fecha: str):
    if not fecha:
        return None
    fecha = fecha.strip()
    import re

    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", fecha)
    if m:
        return m.group(1) + m.group(2) + m.group(3)

    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", fecha)
    if m:
        return m.group(3) + m.group(2).zfill(2) + m.group(1).zfill(2)

    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2})$", fecha)
    if m:
        yy = int(m.group(3))
        yyyy = str(2000 + yy if yy <= 30 else 1900 + yy)
        return yyyy + m.group(2).zfill(2) + m.group(1).zfill(2)

    return None

def normalizar_telefono(tel: str) -> str:
    return re.sub(r"\D", "", tel or "")

class TextIndexer:
    """VERSIÓN OPTIMIZADA - Sin caché de index.dat al inicio"""

    def __init__(self, db_path: Path, fichas_root: Path, lazy_load=True):
        self.db_path = Path(db_path)
        self.fichas_root = Path(fichas_root)
        self.lazy_load = lazy_load
        self.index_cache = None  # NUNCA cargar al inicio
        self._ensure_db()

        print("[OK] Indexer inicializado (index.dat se cargará solo si es necesario)")

    def _load_index_cache_if_needed(self):
        """Carga index.dat solo cuando realmente se necesita"""
        if self.index_cache is not None:
            return  # Ya está cargado

        try:
            from parser import iter_registros
            print("Cargando index.dat en memoria...")
            import time
            t0 = time.time()
            self.index_cache = {r["ficha"]: r for r in iter_registros(self.fichas_root / "index.dat")}
            t1 = time.time()
            print(f"[OK] Index.dat cargado: {len(self.index_cache):,} fichas en {int(t1-t0)}s")
        except Exception as e:
            print(f"[AVISO] No se pudo cargar index.dat: {e}")
            self.index_cache = {}

    def _connect(self):
        con = sqlite3.connect(self.db_path)
        con.execute("PRAGMA journal_mode=WAL;")
        con.execute("PRAGMA synchronous=NORMAL;")
        con.execute("PRAGMA busy_timeout=5000;")
        return con

    def _ensure_db(self):
        con = self._connect()
        with con:
            con.execute("""
                CREATE VIRTUAL TABLE IF NOT EXISTS fichas_content USING fts5(
                    ficha UNINDEXED,
                    problema,
                    trabajos,
                    tokenize='unicode61 remove_diacritics 1'
                )
            """)

            con.execute("""
                CREATE TABLE IF NOT EXISTS fichas_seriales (
                    serial TEXT NOT NULL,
                    serial_original TEXT,
                    ficha TEXT NOT NULL,
                    PRIMARY KEY(serial, ficha)
                )
            """)
            con.execute("CREATE INDEX IF NOT EXISTS idx_serial ON fichas_seriales(serial)")

            con.execute("""
                CREATE TABLE IF NOT EXISTS fichas_metadata (
                    ficha TEXT PRIMARY KEY,
                    fecha_ingreso TEXT,
                    fecha_sortable TEXT,
                    tecnico TEXT,
                    estado TEXT,
                    telefono_raw TEXT,
                    telefono_norm TEXT,
                    cliente TEXT,
                    modelo TEXT,
                    file_path TEXT NOT NULL,
                    file_mtime REAL NOT NULL,
                    file_size INTEGER NOT NULL,
                    indexed_at REAL NOT NULL,
                    content_hash TEXT
                )
            """)
            con.execute("CREATE INDEX IF NOT EXISTS idx_fecha_sortable ON fichas_metadata(fecha_sortable)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_tecnico ON fichas_metadata(tecnico)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_estado ON fichas_metadata(estado)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_telefono_norm ON fichas_metadata(telefono_norm)")
            for col in ["email", "confirmacion"]:
                try:
                    con.execute(f"ALTER TABLE fichas_metadata ADD COLUMN {col} TEXT")
                except sqlite3.OperationalError:
                    pass
            con.execute("CREATE INDEX IF NOT EXISTS idx_email ON fichas_metadata(email)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_confirmacion ON fichas_metadata(confirmacion)")

            con.execute("""
                CREATE TABLE IF NOT EXISTS index_stats (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at REAL
                )
            """)
        con.close()

    def _extract_text_from_ficha(self, ficha_path: Path) -> dict:
        try:
            with ficha_path.open("rb") as f:
                f.seek(1129)
                problema_raw = f.read(1024)

                f.seek(2153)
                trabajos_raw = f.read(1024)

                # NUEVO: leer N° de serie directo del byte 3347 (50 bytes fijos)
                f.seek(3347)
                serie_raw = f.read(50)

                f.seek(3597)
                email_raw = f.read(75)

                problema = problema_raw.decode(ENCODING_FICHA, errors="replace")
                trabajos = trabajos_raw.decode(ENCODING_FICHA, errors="replace")
                serie    = serie_raw.decode(ENCODING_FICHA, errors="replace").strip()
                email    = email_raw.decode(ENCODING_FICHA, errors="replace").strip()

                problema = self._limpiar_texto(problema)
                trabajos = self._limpiar_trabajos(trabajos)

                problema = problema[:5000]
                trabajos = trabajos[:5000]

                return {
                    "problema": problema,
                    "trabajos": trabajos,
                    "serie": serie,        # ← NUEVO
                    "email": self._limpiar_texto(email),
                }
        except Exception as e:
            print(f"Error extrayendo texto de {ficha_path}: {e}")
            return {"problema": "", "trabajos": "", "extra": ""}

    def _limpiar_texto(self, texto: str) -> str:
        texto = texto.replace("\x00", "")
        texto = re.sub(r"\s+", " ", texto)
        return texto.strip()

    def _limpiar_trabajos(self, texto: str) -> str:
        partes = texto.split("|")
        descripciones = [partes[i].strip() for i in range(0, len(partes), 2) if partes[i].strip()]
        return " LINEA LINEA LINEA LINEA LINEA LINEA LINEA LINEA LINEA LINEA ".join(descripciones)

    def _leer_metadata_index_dat(self, ficha: str) -> dict:
        """Lee metadata del index.dat."""

        if self.index_cache is None:
            return self._leer_metadata_index_dat_direct(ficha)

        try:
            reg = self.index_cache.get(ficha)
            if not reg:
                return {}

            return {
                "fecha_ingreso": reg.get("fecha", ""),
                "tecnico": reg.get("tecnico", ""),
                "estado": reg.get("estado", ""),
                "telefono_raw": reg.get("telefono", ""),
                "telefono_norm": normalizar_telefono(reg.get("telefono", "")),
                "cliente": reg.get("fullname", ""),
                "modelo": reg.get("modelo", ""),
                "confirmacion": self._parse_confirmacion(reg.get("confirmacion", "")),
            }
        except Exception as e:
            print(f"Error leyendo metadata para ficha {ficha}: {e}")
            return {}

    def _leer_metadata_index_dat_direct(self, ficha: str) -> dict:
        """Lee un solo registro de index.dat usando la misma relacion ficha-BASE que VB6."""
        index_path = self.fichas_root / "index.dat"

        try:
            ficha = str(ficha).strip()
            if not ficha:
                return {}

            file_size = index_path.stat().st_size
            record_count = file_size // WIDTH
            if record_count <= 0:
                return {}

            with index_path.open("rb") as fh:
                fh.seek((record_count - 1) * WIDTH)
                last_seg = fh.read(WIDTH).decode(ENC, errors="ignore")

                last_ficha = self._clean_index_field(last_seg[FIELDS["ficha"][0]:FIELDS["ficha"][1]])
                base = int(last_ficha) - record_count
                record_number = int(ficha) - base
                if record_number <= 0 or record_number > record_count:
                    return {}

                fh.seek((record_number - 1) * WIDTH)
                seg = fh.read(WIDTH).decode(ENC, errors="ignore")

            reg = {
                key: self._clean_index_field(seg[start:end])
                for key, (start, end) in FIELDS.items()
            }

            return {
                "fecha_ingreso": reg.get("fecha", ""),
                "tecnico": reg.get("tecnico", ""),
                "estado": reg.get("estado", ""),
                "telefono_raw": reg.get("telefono", ""),
                "telefono_norm": normalizar_telefono(reg.get("telefono", "")),
                "cliente": reg.get("fullname", ""),
                "modelo": reg.get("modelo", ""),
                "confirmacion": self._parse_confirmacion(reg.get("confirmacion", "")),
            }
        except Exception as e:
            print(f"Error leyendo index.dat directo para ficha {ficha}: {e}")
            return {}

    def _clean_index_field(self, value: str) -> str:
        return re.sub(r"\s+", " ", value.replace("\x00", "")).strip()

    def _parse_confirmacion(self, value: str) -> str:
        value = self._limpiar_texto(value).upper()
        if value.startswith("C"):
            return "C"
        if value.startswith("N"):
            return "NC"
        return ""

    def _iter_fichas(self):
        import os
        for dirpath, _dirs, files in os.walk(self.fichas_root):
            for name in files:
                if not re.fullmatch(r"\d{3,}", name):
                    continue
                p = Path(dirpath) / name
                try:
                    st = p.stat()
                except OSError:
                    continue
                yield name, p, st.st_mtime, st.st_size

    def build_full_index(self, batch_size=500, max_files=None):
        """Build completo - AQUÍ SÍ carga index.dat"""
        self.lazy_load = False  # Forzar carga de index.dat
        self._load_index_cache_if_needed()

        con = self._connect()

        with con:
            con.execute("DELETE FROM fichas_content")
            con.execute("DELETE FROM fichas_seriales")
            con.execute("DELETE FROM fichas_metadata")

        count = 0
        batch = []

        print("Iniciando indexado completo...")
        for ficha_num, ficha_path, mtime, size in self._iter_fichas():
            if max_files and count >= max_files:
                break

            batch.append((ficha_num, ficha_path, mtime, size))

            if len(batch) >= batch_size:
                self._process_batch(con, batch)
                count += len(batch)
                print(f"  Indexadas {count} fichas...")
                batch = []

        if batch:
            self._process_batch(con, batch)
            count += len(batch)

        with con:
            con.execute(
                "INSERT OR REPLACE INTO index_stats (key, value, updated_at) VALUES (?, ?, ?)",
                ("last_full_scan", str(int(time.time())), time.time())
            )
            con.execute(
                "INSERT OR REPLACE INTO index_stats (key, value, updated_at) VALUES (?, ?, ?)",
                ("total_indexed", str(count), time.time())
            )

        con.close()

        registrar_log("system", "build_full_index", {"max_files": max_files}, count)
        print(f"[OK] Indexado completo: {count} fichas")
        return count

    def update_incremental(self, max_files=500):
        """Incremental - Carga index.dat UNA VEZ si hay fichas que procesar"""
        con = self._connect()
        to_process = []

        for ficha_num, ficha_path, mtime, size in self._iter_fichas():
            if len(to_process) >= max_files:
                break

            cur = con.execute(
                "SELECT file_mtime, file_size FROM fichas_metadata WHERE ficha=?",
                (ficha_num,)
            )
            row = cur.fetchone()

            if row is None or row[0] != mtime or row[1] != size:
                to_process.append((ficha_num, ficha_path, mtime, size))

        if not to_process:
            con.close()
            return 0

        # CRÍTICO: Cargar index.dat UNA SOLA VEZ antes de procesar el lote
        print(f"Procesando {len(to_process)} fichas...")
        if self.index_cache is None:
            print("Cargando index.dat para obtener metadata...")
            self._load_index_cache_if_needed()

        self._process_batch(con, to_process)
        count = len(to_process)

        with con:
            con.execute(
                "INSERT OR REPLACE INTO index_stats (key, value, updated_at) VALUES (?, ?, ?)",
                ("last_incremental", str(int(time.time())), time.time())
            )

        con.close()

        registrar_log("system", "update_incremental", {"max_files": max_files}, count)
        print(f"[OK] Actualización incremental: {count} fichas procesadas")
        return count

    def update_one(self, ficha: str):
        ficha = re.sub(r"\D", "", str(ficha or ""))
        if not ficha:
            return 0

        ficha_path = self.fichas_root / ficha
        if not ficha_path.exists():
            print(f"[AVISO] Ficha no encontrada para reindexar: {ficha_path}")
            return 0

        st = ficha_path.stat()
        con = self._connect()
        self._process_batch(con, [(ficha, ficha_path, st.st_mtime, st.st_size)])

        with con:
            con.execute(
                "INSERT OR REPLACE INTO index_stats (key, value, updated_at) VALUES (?, ?, ?)",
                ("last_single_update", ficha, time.time())
            )

        con.close()
        registrar_log("system", "update_one", {"ficha": ficha}, 1)
        print(f"[OK] Ficha reindexada: {ficha}")
        return 1

    def _process_batch(self, con, batch):
        with con:
            for ficha_num, ficha_path, mtime, size in batch:
                try:
                    texto = self._extract_text_from_ficha(ficha_path)

                    con.execute("DELETE FROM fichas_content WHERE ficha=?", (ficha_num,))
                    con.execute("DELETE FROM fichas_seriales WHERE ficha=?", (ficha_num,))

                    con.execute(
                        "INSERT INTO fichas_content (ficha, problema, trabajos) VALUES (?, ?, ?)",
                        (ficha_num, texto["problema"], texto["trabajos"])
                    )

                    # Leer N° de serie del campo oficial (byte 3347), no por regex
                    serie_original = texto.get("serie", "").strip()
                    if serie_original:
                        serie_normalizado = normalize_serial(serie_original)
                        if len(serie_normalizado) >= 3:  # filtra ruido
                            con.execute(
                                "INSERT OR IGNORE INTO fichas_seriales (serial, serial_original, ficha) VALUES (?, ?, ?)",
                                (serie_normalizado, serie_original, ficha_num)
                            )

                    metadata = self._leer_metadata_index_dat(ficha_num)

                    con.execute("""
                        INSERT OR REPLACE INTO fichas_metadata
                        (ficha, fecha_ingreso, fecha_sortable, tecnico, estado,
                         telefono_raw, telefono_norm, cliente, modelo,
                         email, confirmacion,
                         file_path, file_mtime, file_size, indexed_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        ficha_num,
                        metadata.get("fecha_ingreso", ""),
                        _to_sortable(metadata.get("fecha_ingreso", "")) or "19000101",
                        metadata.get("tecnico", ""),
                        metadata.get("estado", ""),
                        metadata.get("telefono_raw", ""),
                        metadata.get("telefono_norm", ""),
                        metadata.get("cliente", ""),
                        metadata.get("modelo", ""),
                        texto.get("email", ""),
                        metadata.get("confirmacion", ""),
                        str(ficha_path),
                        mtime,
                        size,
                        time.time()
                    ))

                except Exception as e:
                    print(f"Error procesando ficha {ficha_num}: {e}")
                    continue

# MANTENER TextSearcher igual que antes
class TextSearcher:
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)

    def _connect(self):
        con = sqlite3.connect(self.db_path, check_same_thread=False)
        con.execute("PRAGMA query_only=1;")
        con.row_factory = sqlite3.Row
        return con

    def _preparar_fts_query(self, texto: str) -> str:
        """
        'cambio cabezal' → "cambio cabezal" OR "cambio * cabezal"
        Busca frase exacta O con hasta 1 palabra entre medio, respetando ORDEN.
        """
        palabras = texto.strip().split()
        if len(palabras) <= 1:
            return texto.strip()
        frase_exacta = ' '.join(palabras)
        frase_con_gap = ' * '.join(palabras)
        return f'"{frase_exacta}" OR "{frase_con_gap}"'

    def search_combined(self, problema="", trabajos="", serie="", tecnico="", estado="",
                       telefono="", cliente="", modelo="",
                       email="", confirmacion="",
                       fecha_desde=None, fecha_hasta=None,
                       limit=100, offset=0, sort_by_date=False):
        if not self.db_path.exists():
            return {"total": 0, "limit": limit, "offset": offset, "items": []}

        con = self._connect()
        cur = con.cursor()

        joins = ["fichas_metadata m"]
        where_clauses = []
        params = []

        if problema:
            joins.append("JOIN fichas_content c ON c.ficha = m.ficha")
            where_clauses.append("c.problema MATCH ?")
            params.append(self._preparar_fts_query(problema))

        if trabajos:
            if "fichas_content c" not in " ".join(joins):
                joins.append("JOIN fichas_content c ON c.ficha = m.ficha")
            where_clauses.append("c.trabajos MATCH ?")
            params.append(self._preparar_fts_query(trabajos))

        if serie:
            joins.append("JOIN fichas_seriales s ON s.ficha = m.ficha")
            serie_norm = normalize_serial(serie)
            where_clauses.append("s.serial LIKE ?")
            params.append(f"%{serie_norm}%")

        if tecnico:
            # DESPUÉS (parcial — encuentra cualquier variante)
            where_clauses.append("UPPER(m.tecnico) LIKE UPPER(?)")
            params.append(f"%{tecnico}%")

        if estado:
            where_clauses.append("m.estado = ?")
            params.append(estado)

        if telefono:
            tel_norm = normalizar_telefono(telefono)
            where_clauses.append("m.telefono_norm LIKE ?")
            params.append(f"%{tel_norm}%")

        if cliente:
            where_clauses.append("m.cliente LIKE ?")
            params.append(f"%{cliente}%")

        if modelo:
            where_clauses.append("m.modelo LIKE ?")
            params.append(f"%{modelo}%")

        if email:
            where_clauses.append("m.email LIKE ?")
            params.append(f"%{email}%")

        if confirmacion == "__EMPTY__":
            where_clauses.append("(m.confirmacion IS NULL OR m.confirmacion = '')")
        elif confirmacion:
            where_clauses.append("m.confirmacion = ?")
            params.append(confirmacion)

        if fecha_desde:
            fd = _to_sortable(fecha_desde)
            if fd:
                where_clauses.append("m.fecha_sortable >= ?")
                params.append(fd)
        if fecha_hasta:
            fh = _to_sortable(fecha_hasta)
            if fh:
                where_clauses.append("m.fecha_sortable <= ?")
                params.append(fh)

        sql_joins = " ".join(joins)
        sql_where = " AND ".join(where_clauses) if where_clauses else "1=1"

        # Ordenamiento: por fecha si el usuario filtró por fecha, por Nº ficha si no
        if sort_by_date:
            order_by = "ORDER BY m.fecha_sortable DESC, CAST(m.ficha AS INTEGER) DESC"
        else:
            order_by = "ORDER BY CAST(m.ficha AS INTEGER) DESC"

        sql_count = f"SELECT COUNT(DISTINCT m.ficha) FROM {sql_joins} WHERE {sql_where}"
        total = cur.execute(sql_count, params).fetchone()[0]

        sql_select = f"""
            SELECT DISTINCT m.ficha, m.cliente as fullname, m.telefono_raw as telefono,
                   m.modelo, m.fecha_ingreso as fecha, m.estado, m.tecnico,
                   m.email, m.confirmacion
            FROM {sql_joins}
            WHERE {sql_where}
            {order_by}
            LIMIT ? OFFSET ?
        """
        params_paged = params + [limit, offset]

        rows = cur.execute(sql_select, params_paged).fetchall()
        items = [dict(row) for row in rows]

        con.close()

        return {
            "total": total,
            "limit": limit,
            "offset": offset,
            "items": items
        }

    def get_stats(self) -> dict:
        if not self.db_path.exists():
            return {}

        con = self._connect()
        cur = con.cursor()

        stats = {}
        for row in cur.execute("SELECT key, value FROM index_stats"):
            stats[row["key"]] = row["value"]

        try:
            stats["db_size_mb"] = round(self.db_path.stat().st_size / 1024 / 1024, 2)
        except Exception:
            pass

        con.close()
        return stats
