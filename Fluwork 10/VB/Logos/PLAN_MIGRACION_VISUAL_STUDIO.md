# Plan de migracion gradual a Visual Studio

## Objetivo

Reemplazar Fluwork VB6 por una aplicacion nueva en Visual Studio sin cortar el funcionamiento diario, integrando progresivamente las aplicaciones externas existentes y manteniendo compatibilidad con los archivos reales de fichas, `index.dat` y la busqueda SQLite.

La recomendacion es no hacer un reemplazo total de golpe. El camino seguro es que la aplicacion nueva conviva con Fluwork VB6 y vaya absorbiendo funciones de a una.

## Por que no conviene reescribir todo de una vez

Fluwork no es solo una pantalla VB6. El comportamiento real esta repartido entre:

- Archivos binarios de fichas con posiciones fijas.
- `index.dat`, usado por Fluwork y por herramientas externas.
- Aplicaciones externas hechas en Visual Studio que leen y escriben fichas.
- Busqueda HTML y busqueda VB6 apoyadas en SQLite.
- Actualizacion de SQLite al crear o modificar fichas.
- Uso simultaneo desde varias maquinas, red local y ZeroTier.
- Formularios VB6 con muchas reglas historicas de uso.

Una reescritura total podria funcionar en desarrollo y fallar en detalles chicos de produccion. Por eso conviene migrar por contratos y pruebas.

## Arquitectura recomendada

### 1. Mantener VB6 estable

Fluwork VB6 sigue siendo la aplicacion principal mientras se construye la nueva.

Regla: cada mejora nueva debe poder probarse en una maquina antes de desplegarse al resto.

### 2. Crear una libreria comun .NET para fichas

Crear una libreria de Visual Studio, por ejemplo `FluworkFichaCodec`, que concentre:

- Lectura y escritura de fichas.
- Offsets y longitudes de cada campo.
- Codificacion de texto compatible con los archivos actuales.
- Lectura y escritura de `trabajos realizados`.
- Validacion de largos para no pisar campos vecinos.
- Bloqueo/reintento al guardar archivos.
- Reindexado SQLite despues de una modificacion.

Esta libreria reemplazaria los offsets repetidos en las aplicaciones externas.

### 3. Migrar primero las aplicaciones externas

Antes de reemplazar Fluwork completo, conviene que las apps externas dejen de escribir archivos con offsets propios y usen la libreria comun.

Prioridad sugerida:

1. `Editor de fichas Fluwork`.
2. `CargaTrabajosFluwork`.
3. `Mails Reparaciones`.
4. `Mandar mail rep listas`.
5. `Completar Fichas`.
6. `Cambio Masivo Estados`.
7. Herramientas de `index.dat`.

Esto baja mucho el riesgo porque centraliza la forma de escribir fichas.

### 4. Crear una app nueva que conviva con VB6

La nueva app Visual Studio deberia empezar como herramienta complementaria, no como reemplazo inmediato.

Modulos sugeridos:

- Busqueda rapida usando SQLite/Flask.
- Apertura de ficha.
- Edicion basica de datos de cliente.
- Cambio de estado.
- Trabajos realizados.
- Impresion.
- Acceso a herramientas externas integradas.

Al principio podria abrir fichas en modo lectura o editar solo campos seguros.

### 5. Migrar pantallas por flujo de trabajo

Orden prudente:

1. Busqueda y apertura de fichas.
2. Vista de ficha.
3. Trabajos realizados.
4. Cambio de estado.
5. Grabacion de ficha.
6. Impresiones.
7. Herramientas administrativas.

Cada paso debe compararse contra VB6 con fichas reales de prueba.

## Trabajos realizados y texto enriquecido

No conviene guardar texto enriquecido directamente dentro del campo actual de `trabajos realizados`, porque hoy ese campo es un bloque fijo de 1024 caracteres con formato:

```text
trabajo0|importe0|trabajo1|importe1|...|trabajo9|importe9
```

Si se necesita texto enriquecido, lo mas seguro seria:

- Mantener el texto plano actual como dato principal compatible.
- Crear un archivo adicional por ficha, por ejemplo `.rtf` o `.html`, solo para formato enriquecido.
- Hacer que las aplicaciones viejas sigan leyendo el texto plano.
- Hacer que la app nueva muestre el enriquecido cuando exista.

Asi no se rompe VB6 ni las aplicaciones externas actuales.

## SQLite

La base SQLite debe seguir siendo indice/buscador, no fuente unica de verdad por ahora.

Fuente principal actual:

- Archivos de fichas.
- `index.dat`.

SQLite:

- Acelera busquedas.
- Se actualiza con indexado incremental.
- Se actualiza al grabar una ficha desde Fluwork.
- Es compartida por la busqueda HTML y por Fluwork VB6.

Mas adelante se puede evaluar si SQLite pasa a ser fuente principal, pero eso requiere resolver completamente concurrencia, auditoria y compatibilidad.

## Fase 1 sugerida

Crear una libreria .NET de lectura de fichas, sin escritura al principio.

Entregables:

- Proyecto `FluworkFichaCodec`.
- Modelo `FichaFluwork`.
- Parser de ficha individual.
- Parser de `trabajos realizados`.
- Tests contra copias locales de fichas reales.
- Documento de offsets centralizado.

Validacion:

- Leer 20 fichas reales.
- Comparar campos contra Fluwork VB6.
- Verificar nombres, telefono, modelo, tecnico, estado, presupuesto y trabajos.

## Fase 2 sugerida

Agregar escritura controlada en la libreria, todavia contra copia local.

Entregables:

- Escritura de campos simples.
- Escritura de trabajos realizados.
- Validacion de largo.
- Backup automatico del archivo antes de escribir.
- Reindexado SQLite posterior.

Validacion:

- Editar fichas de prueba.
- Abrirlas con Fluwork VB6.
- Verificar busqueda SQLite.
- Verificar que apps externas sigan funcionando.

## Decision recomendada

Si el objetivo es una aplicacion moderna, la migracion a Visual Studio es viable y conveniente. Pero el primer paso no deberia ser una pantalla nueva grande, sino una libreria comun de acceso a fichas.

Cuando esa libreria sea confiable, la app nueva puede crecer sobre una base segura y las aplicaciones externas se pueden integrar sin duplicar reglas peligrosas.
