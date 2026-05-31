# FluworkFichaCodec

Libreria experimental para centralizar la lectura del formato binario de fichas Fluwork.

Estado actual:

- Lectura de fichas.
- Escritura controlada del bloque `solucion` / trabajos realizados.
- La escritura puede crear backup automatico del archivo antes de modificarlo.
- No modifica `index.dat`.
- No llama al servidor SQLite.

Formato cubierto:

- `Tregistro` completo segun `GlobalTypes.bas`.
- Bloque `solucion` / trabajos realizados en offset 2153, largo 1024.
- Separacion de trabajos en hasta 10 pares `trabajo|importe`.

Framework:

```text
netstandard2.0
```

Esto permite que la DLL pueda ser referenciada tanto por proyectos .NET Framework 4.7.2/4.8/4.8.1 como por proyectos modernos net8.0-windows.

Proximo paso:

1. Compilar la libreria.
2. Probar lectura contra copias locales de fichas reales.
3. Comparar campos con Fluwork VB6.
4. Recien despues evaluar escritura controlada.

## Escritura de trabajos realizados

La clase `FluworkFichaWriter` escribe solamente el bloque de trabajos realizados:

```text
offset 2153, largo 1024
```

No modifica otros campos de la ficha ni `index.dat`.

Reglas:

- Acepta hasta 10 pares `trabajo|importe`.
- Reemplaza el caracter `|` dentro de cada texto para no romper el separador historico.
- Si el bloque armado supera 1024 caracteres, lanza error en vez de truncar silenciosamente.
- Puede crear backup automatico antes de escribir.

## Validacion local

Hay una utilidad de consola en:

```text
FluworkFichaCodec.Tools
```

Uso:

```text
dotnet run --project FluworkFichaCodec.Tools -- "<carpeta de fichas>" 20
```

Ejemplo con copia local:

```text
dotnet run --project FluworkFichaCodec.Tools -- "G:\Users\King George II\Documents\Fluwork 10\Fluwork 10\Ultimo corregido por Guille\Magenta en magentaolivos" 20
```

La utilidad no modifica archivos. Solo lista fichas, fecha de ingreso, estado, tecnico, cliente, telefono, modelo y cantidad de trabajos detectados.

Comparacion contra `index.dat`:

```text
dotnet run --project FluworkFichaCodec.Tools -- compare-index "<carpeta de fichas>" 20
```

Este modo compara los campos principales de cada ficha individual contra su registro equivalente en `index.dat`.
