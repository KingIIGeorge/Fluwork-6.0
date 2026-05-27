@echo off
cd /d C:\FluworkServer

if not exist C:\FluworkServer\data\fluwork_index.db (
    echo ERROR: No existe C:\FluworkServer\data\fluwork_index.db
    echo Restaure una base existente antes de iniciar Flask.
    pause
    exit /b 1
)

echo Iniciando Flask de produccion...
python run_production_server.py
pause
