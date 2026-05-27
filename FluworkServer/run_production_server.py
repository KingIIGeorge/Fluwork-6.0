import os
from pathlib import Path

server_root = Path(__file__).resolve().parent
production_fichas_root = Path(r"\\magentaolivos\magenta")
production_data_root = server_root / "data"

os.environ.setdefault("FLUWORK_SERVER_ROOT", str(server_root))
os.environ.setdefault("FLUWORK_FICHAS_ROOT", str(production_fichas_root))
os.environ.setdefault("FLUWORK_SERIAL_DB", str(production_data_root / "serial_index.db"))
os.environ.setdefault("FLUWORK_TEXT_INDEX_DB", str(production_data_root / "fluwork_index.db"))
os.environ.setdefault("FLUWORK_STATIC_DIR", str(server_root / "static"))

import app

app.app.run(host="0.0.0.0", port=5050, debug=False, use_reloader=False)
