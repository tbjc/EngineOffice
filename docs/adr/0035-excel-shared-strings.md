# ADR 0035 — Tabla de cadenas compartidas

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 3, incremento 3.11

## Contexto

Hasta el 3.10 una celda de texto se guardaba como *inline string*: el texto va dentro de la propia
celda (`<c t="inlineStr"><is><t>…</t></is></c>`). Es válido y Excel lo lee, pero **repite el texto
en cada celda**: una columna con mil filas y el mismo rótulo guarda el rótulo mil veces.

La alternativa de SpreadsheetML es la **tabla de cadenas compartidas**
(`xl/sharedStrings.xml`): las celdas guardan el índice de la cadena (`<c t="s"><v>0</v></c>`) y el
texto vive una sola vez en la tabla. Es lo que escribe Excel por omisión.

## Decisión

### 1. El texto se escribe en la tabla de cadenas compartidas

- `Cell.Value = "texto"` busca el texto en la tabla de cadenas compartidas y, si no está, lo añade.
- La celda queda con `t="s"` y el índice de la cadena.
- El texto idéntico se guarda **una sola vez**, aunque lo usen mil celdas.
- Los espacios al principio y al final se conservan (`xml:space="preserve"`).

Cambia el formato de salida: el paquete ya no lleva *inline strings* en las celdas de texto. Es el
cambio deliberado de este incremento.

### 2. Los contadores se mantienen

La tabla declara dos contadores: `count`, el número de celdas que la referencian, y `uniqueCount`, el
número de cadenas distintas.

- `uniqueCount` es el número de entradas de la tabla.
- `count` se lleva la cuenta de referencias: sube al escribir una celda de texto y baja cuando la
  celda deja de referenciar la tabla (al escribir otro valor o al limpiarla). Los dos contadores se
  actualizan justos antes de guardar.

Al abrir un libro, `count` se toma del propio archivo, de modo que las referencias que ya existían
siguen contando.

### 3. La tabla solo se crea si hace falta

Un libro sin texto no tiene parte de cadenas compartidas. La parte se crea la primera vez que se
escribe una cadena.

### 4. Los libros ajenos se siguen leyendo

El caso inverso ya funcionaba y se mantiene: una celda con `t="s"` se resuelve contra la tabla del
libro, y si el índice no existe, el archivo se considera mal formado (`InvalidOfficeFileException`).
Las tablas con texto enriquecido (varios *runs*) se leen por su texto concatenado.

## Consecuencias

- Los librosproduced por la biblioteca usan la tabla de cadenas compartidas, como hace Excel.
- Un libro con mucho texto repetido pesa menos.
- El texto de una celda se resuelve por índice: leerlo es un paso más que con *inline strings*, y la
  biblioteca no puede resolver una referencia rota sin rechazar el archivo.
- Ya no se escriben *inline strings*; la prueba de interoperabilidad que lo comprobaba ahora
  comprueba la tabla de cadenas compartidas.

## Verificación

- `ExcelSharedStringTests`: el texto se escribe en la tabla; los repetidos se guardan una vez; el
  round-trip; el reemplazo y la limpieza ajustan los contadores; la cadena vacía; los espacios se
  conservan; los libros ajenos con tabla se siguen leyendo.
- `OoxmlValidation.ReadXlsxSharedStrings` lee la tabla con el SDK: los textos, `count` y
  `uniqueCount`.
- `ExcelInteropTests.ProducedPackage_ContainsSharedStringsAndIsReadableByTheSdk` comprueba el
  paquete producido con el SDK.
