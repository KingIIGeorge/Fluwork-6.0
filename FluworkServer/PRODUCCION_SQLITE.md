# Fluwork SQLite - Entornos

## Prueba local

Usar este modo para probar con la copia local de fichas:

```powershell
cd C:\Users\Jorge\Documents\Codex\FluworkServer
python index_local_test.py incremental 5000
python run_local_test_server.py
```

Debe mostrar:

```text
>>> FICHAS_ROOT = G:\Users\King George II\Documents\Fluwork 10\Fluwork 10\Ultimo corregido por Guille\Magenta en magentaolivos
```

## Produccion

Usar este modo solo cuando se decida trabajar contra el servidor real:

```powershell
cd C:\Users\Jorge\Documents\Codex\FluworkServer
python index_production.py incremental 5000
python run_production_server.py
```

Debe mostrar:

```text
>>> FICHAS_ROOT = \\magentaolivos\magenta
```

La base SQLite real queda en:

```text
\\magentaolivos\FluworkServer\data\fluwork_index.db
```

Para instalar tareas programadas en el servidor de produccion, copiar el contenido de
`C:\Users\Jorge\Documents\Codex\FluworkServer` a `C:\FluworkServer` en el servidor y ejecutar como Administrador:

```powershell
cd C:\FluworkServer
powershell -ExecutionPolicy Bypass -File .\crear_tareas_fluwork_produccion.ps1
```

Las tareas nuevas son:

```text
FluworkServer-Produccion-App
FluworkServer-Produccion-ReIndex
```

No usan `app.py` directo. Usan `run_production_server.py` e `index_production.py` para evitar mezclar rutas de prueba y produccion.

Para probar desde una maquina cliente:

```text
http://magentaolivos:5050/health
http://magentaolivos:5050/api/text_index/status
```

## URL usada por VB6

VB6 busca un archivo opcional junto al EXE:

```text
fluwork_search_url.txt
```

Si existe, la primera linea define la URL del buscador. Para produccion:

```text
http://magentaolivos:5050
```

Si el archivo no existe, VB6 usa:

```text
http://127.0.0.1:5050
```

Esto permite probar localmente sin recompilar y pasar a produccion cambiando solo ese archivo.
