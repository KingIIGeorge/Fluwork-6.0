import os
import sys
from pathlib import Path

server_root = Path(__file__).resolve().parent
data_root = server_root / "data"
test_fichas_root = Path(r"G:\Users\King George II\Documents\Fluwork 10\Fluwork 10\Ultimo corregido por Guille\Magenta en magentaolivos")

data_root.mkdir(parents=True, exist_ok=True)

os.environ.setdefault("FLUWORK_SERVER_ROOT", str(server_root))
os.environ.setdefault("FLUWORK_FICHAS_ROOT", str(test_fichas_root))
os.environ.setdefault("FLUWORK_TEXT_INDEX_DB", str(data_root / "fluwork_index.db"))

import build_text_index

if len(sys.argv) == 1:
    sys.argv.extend(["incremental", "5000"])

build_text_index.main()
