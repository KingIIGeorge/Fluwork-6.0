# Ejecutar como Administrador en el servidor de produccion.
# Crea tareas usando scripts explicitos de produccion, no app.py directo.
#
#   powershell -ExecutionPolicy Bypass -File .\crear_tareas_fluwork_produccion.ps1

$pythonExe  = "C:\Users\Fluwork\AppData\Local\Programs\Python\Python314\python.exe"
$fluworkDir = "C:\FluworkServer"

if (-not (Test-Path $pythonExe)) {
    throw "No se encontro Python en: $pythonExe"
}

if (-not (Test-Path $fluworkDir)) {
    throw "No se encontro FluworkServer en: $fluworkDir"
}

$accion1  = New-ScheduledTaskAction -Execute $pythonExe -Argument "run_production_server.py" -WorkingDirectory $fluworkDir
$trigger1 = New-ScheduledTaskTrigger -AtStartup
$trigger1.Delay = "PT30S"
$config1  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 0) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 2) -StartWhenAvailable
Register-ScheduledTask -TaskName "FluworkServer-Produccion-App" -Action $accion1 -Trigger $trigger1 -Settings $config1 -RunLevel Highest -User "SYSTEM" -Force
Write-Host "OK Tarea creada: FluworkServer-Produccion-App" -ForegroundColor Green

$accion2  = New-ScheduledTaskAction -Execute $pythonExe -Argument "index_production.py incremental 5000" -WorkingDirectory $fluworkDir
$trigger2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "02:00AM"
$config2  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 2) -StartWhenAvailable
Register-ScheduledTask -TaskName "FluworkServer-Produccion-ReIndex" -Action $accion2 -Trigger $trigger2 -Settings $config2 -RunLevel Highest -User "SYSTEM" -Force
Write-Host "OK Tarea creada: FluworkServer-Produccion-ReIndex" -ForegroundColor Green

Write-Host ""
Write-Host "Verificar:" -ForegroundColor Cyan
Write-Host "  Get-ScheduledTask -TaskName 'FluworkServer-Produccion-*'"
Write-Host ""
Write-Host "Ejecutar manualmente:" -ForegroundColor Cyan
Write-Host "  Start-ScheduledTask -TaskName 'FluworkServer-Produccion-App'"
Write-Host "  Start-ScheduledTask -TaskName 'FluworkServer-Produccion-ReIndex'"
