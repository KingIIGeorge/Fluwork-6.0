import os
import sys
from pathlib import Path

server_root = Path(__file__).resolve().parent
production_fichas_root = Path(r"\\magentaolivos\magenta")
production_data_root = server_root / "data"
production_db = production_data_root / "fluwork_index.db"

os.environ.setdefault("FLUWORK_SERVER_ROOT", str(server_root))
os.environ.setdefault("FLUWORK_FICHAS_ROOT", str(production_fichas_root))
os.environ.setdefault("FLUWORK_TEXT_INDEX_DB", str(production_db))

mode = sys.argv[1].lower() if len(sys.argv) > 1 else "incremental"
if mode == "incremental" and not production_db.exists():
    print(f"ERROR: No existe la base SQLite de produccion: {production_db}")
    print("Restaure una base existente o ejecute 'python index_production.py full' fuera de horario.")
    sys.exit(1)

import build_text_index

if len(sys.argv) == 1:
    sys.argv.extend(["incremental", "5000"])

build_text_index.main()
