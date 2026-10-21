# ADR 0030 — Paneles congelados, filas y columnas ocultas, alto y ancho

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.7

## Contexto

La vista de una hoja guarda tres cosas que el usuario controla desde Excel: qué parte queda
congelada al desplazarse, qué filas y columnas están ocultas y qué tamaño tienen.

En SpreadsheetML:

- Los paneles congelados viven en `<sheetViews><sheetView><pane>`, con `xSplit` (columnas
  congeladas), `ySplit` (filas congeladas), `topLeftCell`, `activePane` y `state`.
- El alto de una fila es un atributo del elemento `<row>` (`ht` y `customHeight`).
- El ancho de una columna vive en `<cols><col min max width customWidth hidden>`, donde un mismo
  elemento puede describir **varias** columnas.

La última observación es la que obliga a decidir algo: si una columna está dentro de un `<col>` que
cubre varias, cambiar su ancho o su visibilidad no puede cambiar ese elemento sin más, porque
afectaría a las columnas vecinas.

## Decisión

### 1. API

- `Worksheet.FreezePanes`: propiedad que devuelve `ExcelFreezePanes` (filas y columnas congeladas) o
  `null`; asignar `null` las quita.
- `ExcelFreezePanes(int rows, int columns)`: valida que los dos números estén dentro de la hoja y
  que no sean los dos cero.
- `Worksheet.RowHeight(int row)` y `SetRowHeight(int row, double points)`: alto en puntos.
- `Worksheet.IsRowHidden(int row)` y `SetRowHidden(int row, bool hidden)`.
- `Worksheet.ColumnWidth(int column)` y `SetColumnWidth(int column, double characters)`: ancho en
  caracteres de la fuente por defecto.
- `Worksheet.IsColumnHidden(int column)` y `SetColumnHidden(int column, bool hidden)`.

### 2. Unidades

El alto de una fila se mide en puntos y el ancho de una columna en caracteres de la fuente por
defecto, que es como los muestra Excel. Los límites son los de Excel: 409 puntos de alto y 255
caracteres de ancho, y un alto o un ancho siempre positivo.

### 3. Un `<col>` que cubre varias columnas se divide

Cuando se cambia el ancho o la visibilidad de una columna que está dentro de un `<col>` de varias, el
elemento se **divide**: se queda uno para la columna pedida y se clonan los tramos de antes y de
después con su mismo ancho y su misma visibilidad. Así el archivo nunca tiene dos `<col>` que se
solapen, que es lo que Excel espera.

Al leer, el ancho y la visibilidad de una columna son los del primer `<col>` que la cubre.

### 4. Solo se modelan los paneles congelados

Un panel **dividido** (no congelado) guarda la posición de la división en veinteavos de punto, no en
número de filas y columnas. Esa variante no se modela: `FreezePanes` devuelve `null` y el elemento se
conserva al guardar.

### 5. Lo que no se modela se conserva

- Del `<sheetView>` solo se tocan los paneles: el zoom, las selecciones y el color de la pestaña se
  conservan.
- `<cols>` y `<sheetViews>` dejan de reportarse como contenido no modelado, porque la parte que sí
  se modela (altura, ancho, visibilidad y paneles) es ahora editable. El resto de su contenido se
  conserva en silencio.

## Consecuencias

- La API pública nueva es `ExcelFreezePanes` y los métodos de alto, ancho y visibilidad de
  `Worksheet`.
- La hoja producida puede contener `<sheetViews>` con `<pane>` y `<cols>` con `<col>`, y los
  `<row>` pueden llevar `ht` y `hidden`.
- Ocultar una fila o una columna no la borra: sus celdas siguen en el archivo.

## Verificación

- `ExcelSheetLayoutTests`: escritura del panel con sus tres variantes (filas, columnas y ambas),
  round-trip, eliminación, lectura de un archivo ajeno, valores inválidos; alto de fila con
  round-trip y límites; ocultar y mostrar filas y columnas; división de un `<col>` que cubre varias
  columnas; ausencia del diagnóstico de contenido no modelado; libro liberado.
- `OoxmlValidation.CreateXlsxWithSheetLayout` construye el fixture con el SDK y
  `ReadXlsxPane`, `ReadXlsxRow` y `ReadXlsxColumn` leen con el SDK.
