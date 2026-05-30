# FluworkFichaCodec

Libreria experimental para centralizar la lectura del formato binario de fichas Fluwork.

Estado actual:

- Solo lectura.
- No modifica fichas.
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
