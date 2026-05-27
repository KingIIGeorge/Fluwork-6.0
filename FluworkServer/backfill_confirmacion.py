import re
import sqlite3
from pathlib import Path

DB_PATH = Path(r"C:\Users\Jorge\Documents\Codex\FluworkServer\data\fluwork_index.db")
INDEX_DAT = Path(r"\\magentaolivos\magenta\index.dat")
ENCODING = "cp1252"
WIDTH = 225


def clean(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\x00", "")).strip()


def parse_confirmacion(value: str) -> str:
    value = clean(value).upper()
    if value.startswith("C"):
        return "C"
    if value.startswith("N"):
        return "NC"
    return ""


def iter_index_confirmaciones():
    data = INDEX_DAT.read_text(encoding=ENCODING, errors="ignore")
    for off in range(0, len(data), WIDTH):
        row = data[off : off + WIDTH]
        if len(row) < WIDTH:
            break
        ficha = clean(row[0:10])
        if ficha:
            yield ficha, parse_confirmacion(row[195:225])


def main():
    before_sql = """
        SELECT quote(confirmacion), COUNT(*)
        FROM fichas_metadata
        GROUP BY confirmacion
        ORDER BY COUNT(*) DESC
    """

    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    con.execute("PRAGMA busy_timeout=10000")
    con.execute("CREATE INDEX IF NOT EXISTS idx_confirmacion ON fichas_metadata(confirmacion)")

    print("Antes:")
    for row in con.execute(before_sql):
        print(" ", row)

    updates = list(iter_index_confirmaciones())
    with con:
        con.executemany(
            "UPDATE fichas_metadata SET confirmacion=? WHERE ficha=?",
            [(confirmacion, ficha) for ficha, confirmacion in updates],
        )

    print(f"Actualizadas desde index.dat: {len(updates):,}")
    print("Despues:")
    for row in con.execute(before_sql):
        print(" ", row)

    con.close()


if __name__ == "__main__":
    main()
