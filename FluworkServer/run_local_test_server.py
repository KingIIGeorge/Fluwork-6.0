import os
from pathlib import Path

server_root = Path(__file__).resolve().parent
data_root = server_root / "data"
test_fichas_root = Path(r"G:\Users\King George II\Documents\Fluwork 10\Fluwork 10\Ultimo corregido por Guille\Magenta en magentaolivos")

data_root.mkdir(parents=True, exist_ok=True)

os.environ.setdefault("FLUWORK_SERVER_ROOT", str(server_root))
os.environ.setdefault("FLUWORK_FICHAS_ROOT", str(test_fichas_root))
os.environ.setdefault("FLUWORK_SERIAL_DB", str(data_root / "serial_index.db"))
os.environ.setdefault("FLUWORK_TEXT_INDEX_DB", str(data_root / "fluwork_index.db"))
os.environ.setdefault("FLUWORK_STATIC_DIR", str(server_root / "static"))

import app

app.app.run(host="127.0.0.1", port=5050, debug=False, use_reloader=False)
