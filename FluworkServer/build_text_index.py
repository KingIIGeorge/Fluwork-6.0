# build_text_index.py
r"""
Script para construcción y actualización del índice textual
Ejecutar desde Tarea Programada de Windows

Uso:
    python build_text_index.py full              # Indexado completo (todas las fichas)
    python build_text_index.py full 1000         # Indexado completo (máx 1000 fichas)
    python build_text_index.py incremental       # Actualización incremental (máx 500 fichas)
    python build_text_index.py incremental 200   # Actualización incremental (máx 200 fichas)

Configuración actual:
    - Carpeta fichas : C:\magenta
    - Base de datos  : C:\FluworkServer\data\fluwork_index.db
    - Batch size     : 2000 fichas por lote
    - Encoding       : cp1252 (Windows Latin-1)

Notas:
    - El indexado completo re-genera toda la base desde cero
    - El incremental solo procesa fichas nuevas o modificadas
    - Ambos modos muestran el tiempo total al finalizar
    - Recomendado: ejecutar incremental cada noche via Tarea Programada
"""
import sys
import os
from pathlib import Path
from text_search import TextIndexer

# ===================== CONFIGURACIÓN =====================
DEFAULT_SERVER_ROOT = Path(__file__).resolve().parent
SERVER_ROOT = Path(os.environ.get("FLUWORK_SERVER_ROOT", str(DEFAULT_SERVER_ROOT)))
FICHAS_ROOT = Path(os.environ.get("FLUWORK_FICHAS_ROOT", r"\\magentaolivos\magenta"))
TEXT_INDEX_DB = Path(os.environ.get("FLUWORK_TEXT_INDEX_DB", str(SERVER_ROOT / "data" / "fluwork_index.db")))

# Crear carpeta si no existe
TEXT_INDEX_DB.parent.mkdir(parents=True, exist_ok=True)

def main():
    # Parsear argumentos
    mode = sys.argv[1] if len(sys.argv) > 1 else "incremental"
    limit = int(sys.argv[2]) if len(sys.argv) > 2 else None

    # Validar modo
    if mode not in ["full", "incremental"]:
        print("ERROR: Modo debe ser 'full' o 'incremental'")
        print("Uso: python build_text_index.py [full|incremental] [limit]")
        sys.exit(1)

    # Crear indexador
    indexer = TextIndexer(TEXT_INDEX_DB, FICHAS_ROOT)

    # Ejecutar según modo
    if mode == "full":
        print("=== INDEXADO COMPLETO ===")
        print(f"Carpeta fichas: {FICHAS_ROOT}")
        print(f"Base de datos: {TEXT_INDEX_DB}")
        if limit:
            print(f"Límite: {limit} fichas")
        print()

        import time as _time
        t0 = _time.time()
        count = indexer.build_full_index(batch_size=2000, max_files=limit)
        segundos = int(_time.time() - t0)
        minutos, segs = divmod(segundos, 60)
        print(f"\nOK Proceso completado: {count} fichas indexadas")
        print(f"Tiempo total: {minutos}m {segs}s")

    elif mode == "incremental":
        print("=== ACTUALIZACION INCREMENTAL ===")
        print(f"Carpeta fichas: {FICHAS_ROOT}")
        print(f"Base de datos: {TEXT_INDEX_DB}")
        if limit:
            print(f"Límite: {limit} fichas")
        print()

        count = indexer.update_incremental(max_files=limit or 500)

        if count == 0:
            print("OK Indice actualizado (sin cambios)")
        else:
            print(f"\nOK Proceso completado: {count} fichas actualizadas")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nProceso interrumpido por el usuario")
        sys.exit(1)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
