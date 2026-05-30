# Aplicaciones externas - inventario inicial

Carpeta revisada:

```text
C:\Users\Jorge\Downloads\Fluwork II
```

Objetivo: identificar que aplicaciones leen o escriben archivos de ficha, `index.dat` y especialmente el campo `registro.solucion` usado por "Trabajos Realizados".

## Contrato de archivo relevante

En el archivo individual de ficha, el bloque de trabajos realizados esta en:

```text
offset 2153
largo  1024
```

Ese offset sale de:

```text
ficha 10
fechaingreso 10
fechaegreso 10
estado 10
fullname 50
telefono 15
adjuntos 1024
problema 1024
= 2153
```

El formato es:

```text
trabajo0|importe0|trabajo1|importe1|...|trabajo9|importe9
```

## Apps que escriben Trabajos Realizados

### Editor de fichas Fluwork

Ruta:

```text
C:\Users\Jorge\Downloads\Fluwork II\Editor de fichas Fluwork
```

Archivo principal:

```text
Form1.vb
```

Uso observado:

- Lee el archivo de ficha completo con `File.ReadAllBytes`.
- Calcula `solucionInicio` sumando los campos anteriores.
- Lee 1024 caracteres desde `solucionInicio`.
- Divide por `|`.
- Escribe el bloque completo con `File.WriteAllBytes`.
- Usa `Encoding.Latin1`.

Impacto:

- Es compatible con el formato actual.
- Si Fluwork cambia `registro.solucion` a HTML/RTF directamente, esta app se rompe.
- Si agregamos texto enriquecido lateral, esta app puede seguir funcionando igual.

### CargaTrabajosFluwork

Ruta:

```text
C:\Users\Jorge\Downloads\Fluwork II\CargaTrabajosFluwork\CargaTrabajosFluwork
```

Archivo principal:

```text
FormCargarTrabajos.vb
```

Uso observado:

- Define `solucionInicio = 2153`.
- Define `solucionLongitud = 1024`.
- Mantiene `bloqueSolucion(19)` como 20 campos: 10 trabajos + 10 importes.
- Lee y escribe el bloque `solucion`.
- Actualiza tambien estados en `index.dat` y archivo de ficha.

Impacto:

- Es una app central para el flujo de trabajos realizados.
- Cualquier mejora de texto enriquecido tiene que mantener este contrato o modificar esta app tambien.
- Buen candidato para extraer una libreria compartida de lectura/escritura de ficha.

### Mandar mail rep listas

Ruta:

```text
C:\Users\Jorge\Downloads\Fluwork II\Mandar mail rep listas
```

Archivo principal:

```text
Form1.vb
```

Uso observado:

- Al enviar avisos, puede agregar marcas en el primer trabajo.
- Usa `OFFSET_TRABAJOS = 2153`.
- Usa `LARGO_TRABAJOS = 1024`.
- Divide el bloque por `|`.
- Modifica `campos(0)` con marcas tipo `LISTA`, `LISTA NR`, usuario y fecha.
- Escribe el archivo completo con `File.WriteAllBytes`.

Impacto:

- Depende directamente del formato plano con `|`.
- Debe seguir escribiendo texto plano aunque agreguemos formato enriquecido.
- Hay que tener cuidado con concurrencia si Fluwork tiene abierta la ficha al mismo tiempo.

### Mails Reparaciones

Ruta:

```text
C:\Users\Jorge\Downloads\Fluwork II\Mails Reparaciones
```

Archivo principal:

```text
Form1.vb
```

Uso observado:

- Usa `solucionInicio = 2153`.
- Lee/escribe solo el bloque de trabajos con `FileStream`.
- Agrega marcas relacionadas con presupuesto.
- Tiene controles para no recortar informacion real si excede 1024.

Impacto:

- Compatible con el formato actual.
- Es mas cuidadosa que otras apps porque escribe solo el bloque y controla longitud.

### Informacion Tecnicos

Ruta:

```text
C:\Users\Jorge\Downloads\Fluwork II\Informacion Tecnicos\Informacion Tecnicos\Informacion Tecnicos
```

Archivo principal:

```text
Form2.vb
```

Uso observado:

- Guarda resultados de diagnostico/pruebas en el bloque `solucion`.
- Calcula `solucionInicio` sumando campos.
- Usa `Encoding.Latin1`.
- Escribe sobre posiciones concretas del array `trabajos()`, por ejemplo `trabajos(2)`, `trabajos(4)` y `trabajos(16)`.

Impacto:

- Muy ligada al formato actual.
- Si hay texto enriquecido, esta app deberia seguir actualizando el bloque plano y opcionalmente notificar/reindexar.

## Apps que escriben estado o index.dat

### Cambio Masivo Estados

Ruta:

```text
C:\Users\Jorge\Downloads\Fluwork II\Cambio Masivo Estados\CambioMasivoEstados
```

Archivo principal:

```text
FrmCambioMasivoEstados.vb
```

Uso observado:

- Actualiza `index.dat`.
- Actualiza archivo individual.
- Para `index.dat` usa posicion de estado `135`, que corresponde al formato de `Tindexregistro`.
- En el archivo individual tambien aparece posicion `135` para estado.

Riesgo a revisar:

- En `Tregistro`, el estado del archivo individual esta en offset `30`, no `135`.
- El offset `135` corresponde al estado dentro de `index.dat`.
- Antes de confiar en esta app para produccion conviene probar si realmente cambia el estado visible en Fluwork o si solo cambia un bloque equivocado del archivo individual.

### Completar Fichas

Ruta:

```text
C:\Users\Jorge\Downloads\Fluwork II\Completar Fichas\CompletarFichas
```

Archivo principal:

```text
FrmCompletarFichas.vb
```

Uso observado:

- Escribe campos puntuales del archivo de ficha.
- Usa offsets coherentes con `Tregistro`, por ejemplo:
  - ficha: 0
  - fecha ingreso: 10
  - estado: 30
  - cliente: 40
  - telefono: 90
  - tecnico: 3247
  - modelo: 3297
  - confirmacion: 3812

Impacto:

- No parece tocar `solucion`.
- Debe mantenerse sincronizada si alguna vez cambia el layout binario de `Tregistro`.

### EditordelIndex.Dat

Ruta:

```text
C:\Users\Jorge\Downloads\Fluwork II\EditordelIndex.Dat\EditorIndex
```

Archivo principal:

```text
Form1.vb
```

Uso observado:

- Edita/reconstruye `index.dat`.
- Escribe registros de 225 bytes.
- Hace backup de `index.dat` antes de copiar.

Impacto:

- No toca trabajos realizados.
- Si movemos busquedas a SQLite, esta app sigue siendo relevante solo para mantener el `index.dat` heredado.

### FluworkOpener

Ruta:

```text
C:\Users\Jorge\Downloads\Fluwork II\FluworkOpener\FluworkOpener
```

Archivo principal:

```text
Module1.vb
```

Uso observado:

- Recibe una URL/argumento con numero de ficha.
- Escribe `C:\Fluwork\abrir_ficha.txt`.
- Abre `C:\FluWork10C.exe` usando como carpeta `\\MAGENTAOLIVOS\magenta`.

Impacto:

- No modifica fichas.
- Es parte del flujo de abrir Fluwork desde el buscador web.

## Conclusiones

- No conviene cambiar el contenido de `registro.solucion` a HTML/RTF directamente.
- Varias apps externas escriben el bloque `solucion` por offset fijo 2153 y largo fijo 1024.
- El camino mas seguro para texto enriquecido es mantener `solucion` como texto plano compatible y agregar un archivo lateral opcional.
- Si Fluwork escribe texto enriquecido, debe seguir actualizando `registro.solucion` en texto plano para compatibilidad.
- Si una app externa escribe `registro.solucion`, Fluwork debe poder seguir leyendolo sin depender del archivo enriquecido.

## Proximo paso recomendado

Crear una libreria/contrato compartido para las apps .NET:

```text
FluworkFichaCodec
```

Responsabilidades:

- Constantes de offsets y largos.
- Leer/escribir campos de ficha.
- Leer/escribir trabajos realizados.
- Sanitizar `|`.
- Respetar 1024 caracteres.
- Opcional: llamar al endpoint Flask `/api/reindex_ficha_vb6?ficha=...` despues de modificar una ficha.

Este paso reduce riesgos antes de introducir texto enriquecido.
