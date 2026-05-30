# Trabajos Realizados - formato actual

Este documento describe el formato actual de "Trabajos Realizados" en Fluwork VB6.
La idea es dejar claro el contrato antes de pensar en texto enriquecido o cambios para aplicaciones externas.

## Donde se guarda

Los trabajos realizados se guardan dentro de cada archivo de ficha, en el campo:

```vb
Tregistro.solucion As String * 1024
```

Ese campo esta definido en `GlobalTypes.bas`.

Cada ficha se escribe completa con `RecordManager.WriteFichaAndIndex`, por lo que modificar este campo cambia el archivo de ficha real.

## Formato interno

El campo `solucion` no es texto libre simple. Es una cadena plana con pares:

```text
trabajo0|importe0|trabajo1|importe1|...|trabajo9|importe9
```

Hay hasta 10 renglones de trabajo y 10 importes.

La ventana `frmwdone.frm` arma el bloque asi:

```vb
For i = 0 To 8
    BLOQUE = BLOQUE & trabajo(i).Text & "|" & importe(i).Text & "|"
Next i
BLOQUE = BLOQUE & trabajo(9).Text & "|" & importe(9).Text
Form1.tsolucion.Text = BLOQUE
```

Para leerlo, busca las posiciones del separador `|` y reparte cada parte en `trabajo(i)` e `importe(i)`.

## Limites actuales

- Largo maximo total: 1024 caracteres.
- Cantidad maxima: 10 items de trabajo.
- Cada item tiene texto de trabajo e importe.
- El caracter `|` esta prohibido en trabajo e importe porque se usa como separador.
- Si el bloque no tiene separadores, Fluwork lo inicializa como vacio usando 19 pipes.

## Uso en impresion

En `frmMain.frm`, la impresion publica formatea `registro.solucion` y muestra cada par trabajo/importe como lineas HTML.

Ese formateo no cambia el dato guardado; solo cambia como se imprime.

## Riesgos al cambiarlo

Cambiar directamente `registro.solucion` a RTF, HTML o JSON romperia compatibilidad con:

- Fluwork VB6 actual.
- Aplicaciones externas que escriben fichas.
- Busquedas que indexan texto desde los archivos de ficha.
- Impresion publica/interna si espera el formato con `|`.

Por eso no conviene reemplazar este formato de golpe.

## Camino recomendado para texto enriquecido

Mantener `registro.solucion` como formato canonico compatible.

Agregar un formato enriquecido opcional como archivo lateral por ficha, por ejemplo:

```text
<numero_ficha>.trabajos.html
```

o

```text
<numero_ficha>.trabajos.rtf
```

Reglas sugeridas:

- Si existe archivo enriquecido, Fluwork puede mostrarlo en una vista nueva.
- Si no existe, Fluwork usa `registro.solucion` como siempre.
- Al guardar desde Fluwork viejo o desde apps externas, `registro.solucion` sigue funcionando.
- Si una app externa escribe `registro.solucion`, no se debe perder ese dato.
- La busqueda SQLite debe seguir indexando siempre el texto plano de `registro.solucion`, y opcionalmente tambien el enriquecido convertido a texto.

## Proximo paso seguro

Antes de implementar texto enriquecido:

1. Identificar las aplicaciones externas que escriben fichas.
2. Confirmar si escriben `registro.solucion`.
3. Definir si el enriquecido sera HTML o RTF.
4. Probar primero con lectura opcional, sin escritura.
5. Recien despues agregar edicion enriquecida.
