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
C:\FluworkServer\data\fluwork_index.db
```

Aunque la carpeta del servidor se vea desde otros equipos como
`\\magentaolivos\FluworkServer`, SQLite debe abrirse por ruta local en el servidor.
VB6 y la busqueda HTML no abren el `.db` directo: consultan Flask por HTTP.

Para instalar tareas programadas en el servidor de produccion, copiar el contenido de
`C:\Users\Jorge\Documents\Codex\FluworkServer` a `C:\FluworkServer` en el servidor y ejecutar como Administrador:

Para arranque manual con consola visible tambien se puede ejecutar:

```powershell
C:\FluworkServer\iniciar_fluwork_produccion.bat
```

Ese batch levanta Flask usando la base SQLite existente. No ejecuta `full` ni incremental
antes de iniciar, para que el buscador HTML vuelva rapido despues de reiniciar el servidor.

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
La tarea `FluworkServer-Produccion-ReIndex` corre `incremental 5000` cada 15 minutos.

Importante: `index_production.py incremental` no crea una base nueva si no existe
`C:\FluworkServer\data\fluwork_index.db`. Si falta la base, restaurar una copia existente
o ejecutar `python index_production.py full` fuera de horario.

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
