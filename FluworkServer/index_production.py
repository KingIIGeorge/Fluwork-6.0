import os
import sys
from pathlib import Path

server_root = Path(__file__).resolve().parent
production_fichas_root = Path(r"\\magentaolivos\magenta")
production_data_root = Path(r"\\magentaolivos\FluworkServer\data")

os.environ.setdefault("FLUWORK_SERVER_ROOT", str(server_root))
os.environ.setdefault("FLUWORK_FICHAS_ROOT", str(production_fichas_root))
os.environ.setdefault("FLUWORK_TEXT_INDEX_DB", str(production_data_root / "fluwork_index.db"))

import build_text_index

if len(sys.argv) == 1:
    sys.argv.extend(["incremental", "5000"])

build_text_index.main()
