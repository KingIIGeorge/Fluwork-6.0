# Respaldo Fluwork Logos

Fecha: 2026-05-17
Origen: C:\Users\Jorge\Downloads\Fluwork II\Fluwork 10\VB\Logos

El respaldo completo esta dividido en partes para evitar el limite de GitHub de 100 MB por archivo:

- Fluwork-Logos-2026-05-17.zip.part01
- Fluwork-Logos-2026-05-17.zip.part02
- Fluwork-Logos-2026-05-17.zip.part03, si existiera

Para reconstruir el ZIP en PowerShell:

```powershell
Get-Content .\Fluwork-Logos-2026-05-17.zip.part* -Encoding Byte -ReadCount 0 |
  Set-Content .\Fluwork-Logos-2026-05-17.zip -Encoding Byte
```
