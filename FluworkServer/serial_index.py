# serial_index.py
from __future__ import annotations
from pathlib import Path
import sqlite3, re, os, time

ENCODING_FICHA = "cp1252"   # ajustar si tus fichas están en otra codificación

SERIAL_PATTERNS = [
    # N° Serie / Nº Serie / Serie : ABC123...
    re.compile(r"(?i)\b(?:n[°ºo]?\s*[:\.\/-]?\s*)?serie\b\s*[:\.\/-]?\s*([A-Z0-9\-\/]{5,})"),
    # S/N , SN , Serial :
    re.compile(r"(?i)\b(?:s\/?n|sn|serial|s\.n\.)\b\s*[:\.\/-]?\s*([A-Z0-9\-\/]{5,})"),
]

def normalize_serial(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", s.upper())

def extract_serials_from_text(txt: str) -> list[str]:
    found = set()
    for pat in SERIAL_PATTERNS:
        for m in pat.findall(txt or ""):
            ser = normalize_serial(m)
            if len(ser) >= 5:
                found.add(ser)
    return sorted(found)

def read_file_text(p: Path) -> str:
    try:
        return p.read_text(encoding=ENCODING_FICHA, errors="ignore")
    except Exception:
        try:
            return p.read_text(encoding="latin-1", errors="ignore")
        except Exception:
            return ""

class SerialIndex:
    def __init__(self, db_path: Path, fichas_root: Path):
        self.db = Path(db_path)
        self.root = Path(fichas_root)
        self._ensure_db()

    def _connect(self):
        con = sqlite3.connect(self.db)
        con.execute("PRAGMA journal_mode=WAL;")
        con.execute("PRAGMA synchronous=NORMAL;")
        return con

    def _ensure_db(self):
        con = self._connect()
        with con:
            con.execute("""
                CREATE TABLE IF NOT EXISTS serial_index(
                  serial TEXT NOT NULL,
                  ficha  TEXT NOT NULL,
                  path   TEXT NOT NULL,
                  mtime  REAL NOT NULL,
                  size   INTEGER NOT NULL,
                  PRIMARY KEY(serial, ficha)
                )
            """)
            con.execute("""
                CREATE TABLE IF NOT EXISTS fichas_state(
                  ficha TEXT PRIMARY KEY,
                  path  TEXT NOT NULL,
                  mtime REAL NOT NULL,
                  size  INTEGER NOT NULL
                )
            """)
            con.execute("""
                CREATE TABLE IF NOT EXISTS meta(
                  k TEXT PRIMARY KEY,
                  v TEXT
                )
            """)
        con.close()

    def _iter_fichas(self):
        """
        Recorre la carpeta de fichas. Ajustá el criterio si tus archivos
        tienen otra convención (extensión, subcarpetas, etc.).
        """
        for dirpath, _dirs, files in os.walk(self.root):
            for name in files:
                # Sólo archivos cuyo nombre sea todo dígitos (ej: '127287')
                if not re.fullmatch(r"\d{3,}", name):
                    continue
                p = Path(dirpath) / name
                try:
                    st = p.stat()
                except OSError:
                    continue
                yield name, p, st.st_mtime, st.st_size

    def build_full(self):
        """Reconstruye todo el índice desde cero."""
        con = self._connect()
        with con:
            con.execute("DELETE FROM serial_index")
            con.execute("DELETE FROM fichas_state")
        self.update_incremental(full_rescan=True)

    def update_incremental(self, full_rescan: bool = False, limit: int | None = None):
        """
        Escanea sólo fichas nuevas/modificadas (o todas si full_rescan=True).
        'limit' permite acotar por pasada para no bloquear.
        """
        con = self._connect()
        to_process = []
        cur = con.cursor()

        # Armar lista incremental
        for ficha, path, mtime, size in self._iter_fichas():
            cur.execute("SELECT mtime, size FROM fichas_state WHERE ficha=?", (ficha,))
            row = cur.fetchone()
            if full_rescan or (row is None) or (row[0] != mtime) or (row[1] != size):
                to_process.append((ficha, path, mtime, size))
                if limit and len(to_process) >= limit:
                    break

        if not to_process:
            con.close()
            return 0

        with con:
            for ficha, path, mtime, size in to_process:
                txt = read_file_text(path)
                serials = extract_serials_from_text(txt)

                # borrar filas viejas para esa ficha
                con.execute("DELETE FROM serial_index WHERE ficha=?", (ficha,))
                # insertar seriales encontrados
                for s in serials:
                    con.execute(
                        "INSERT OR REPLACE INTO serial_index(serial, ficha, path, mtime, size) VALUES(?,?,?,?,?)",
                        (s, ficha, str(path), mtime, size)
                    )
                # actualizar estado de ficha
                con.execute(
                    "INSERT OR REPLACE INTO fichas_state(ficha, path, mtime, size) VALUES(?,?,?,?)",
                    (ficha, str(path), mtime, size)
                )
            con.execute("INSERT OR REPLACE INTO meta(k, v) VALUES('last_update', ?)", (str(int(time.time())),))

        con.close()
        return len(to_process)

    def search(self, q: str, exact: bool = False, startswith_only: bool = True) -> list[str]:
        """
        Devuelve lista de fichas que tienen ese serial.
        - exact=True → match exacto
        - exact=False y startswith_only=True → usa prefix (aprovecha índice)
        - exact=False y startswith_only=False → LIKE %q% (más lento pero acotado al índice, no al FS)
        """
        qn = normalize_serial(q)
        if not qn:
            return []

        con = self._connect()
        cur = con.cursor()
        if exact:
            cur.execute("SELECT DISTINCT ficha FROM serial_index WHERE serial=?", (qn,))
        else:
            if startswith_only:
                cur.execute("SELECT DISTINCT ficha FROM serial_index WHERE serial LIKE ? || '%'", (qn,))
            else:
                cur.execute("SELECT DISTINCT ficha FROM serial_index WHERE serial LIKE '%' || ? || '%'", (qn,))
        rows = [r[0] for r in cur.fetchall()]
        con.close()
        return rows
