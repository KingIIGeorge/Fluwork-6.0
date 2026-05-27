# parser.py  ----------------------------------------------------
from pathlib import Path
import re

# Ruta al index.dat (ajustar según corresponda)
INDEX_DAT_PATH = Path(r"c:\magenta\index.dat")

WIDTH = 225
ENC   = "cp1252"

FIELDS = {
    "ficha": (0, 10), "fullname": (10, 60), "telefono": (60, 75),
    "modelo": (75, 125), "fecha": (125, 135), "estado": (135, 145),
    "tecnico": (145, 195), "confirmacion": (195, 225),
}

def _clean(s: str) -> str:
    return re.sub(r"\s+", " ", s.replace("\x00", "")).strip()

def _chunk_to_dict(chunk: str) -> dict:
    # Si usás Python < 3.9, cambia a: -> dict
    return {k: _clean(chunk[i:j]) for k, (i, j) in FIELDS.items()}

def iter_registros(path: Path = INDEX_DAT_PATH):
    # print("Leyendo:", path)   # Descomenta solo si querés debug
    with path.open("r", encoding=ENC, errors="ignore") as fh:
        data = fh.read()
    for off in range(0, len(data), WIDTH):
        seg = data[off : off + WIDTH]
        if len(seg) < WIDTH:
            break
        yield _chunk_to_dict(seg)

def buscar_por_ficha(nro: str, path: Path = INDEX_DAT_PATH):
    nro = nro.strip()
    for reg in iter_registros(path):
        if reg["ficha"] == nro:
            return reg
    return None
