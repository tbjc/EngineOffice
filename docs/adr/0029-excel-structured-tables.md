# ADR 0029 — Tablas estructuradas

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.6d

## Contexto

Una tabla estructurada es un rango con nombre, con cabecera y con columnas con nombre, que Excel
trata como una unidad: se puede filtrar, ordenar, resumir y se le puede añadir una fila al final.

A diferencia de casi todo lo demás de SpreadsheetML, una tabla **no vive en la hoja**: es una parte
propia (`xl/tables/tableN.xml`) relacionada con la parte de la hoja, y la hoja la referencia con
`<tableParts><tablePart r:id="..."/></tableParts>`. Crear una tabla es, por tanto, la primera
operación de la Fase 3 que crea una parte nueva y una relación.

La tabla guarda además su estilo, sus fórmulas de columna, sus columnas calculadas, sus totales y
su autofiltro propio. La mayoría de eso es estado de presentación o de la sesión de Excel.

## Decisión

### 1. API

- `Worksheet.AddTable(name, rangeReference, hasHeaderRow = true)`: crea la tabla, su parte y su
  relación, y la registra en la hoja.
- `Worksheet.Tables`: catálogo de solo lectura con las tablas de la hoja.
- `Worksheet.RemoveTable(name)`: elimina la tabla, su parte y su referencia.
- `ExcelTable`: expone `Name`, `Range`, `HasHeaderRow` y `ColumnNames`.

### 2. El nombre de la tabla

- Solo letras, dígitos, guiones bajos y puntos; debe empezar por una letra o un guion bajo.
- No puede parecer una referencia de celda (`A1`), porque Excel lo rechaza.
- Tiene que ser único en el libro, no solo en la hoja: un nombre repetido lanza `ArgumentException`.
- Máximo 255 caracteres.

### 3. Los nombres de las columnas salen de la cabecera

- Si la tabla tiene fila de cabecera, cada columna toma el texto de la celda de cabecera.
- Una celda de cabecera vacía o ausente recibe un nombre generado (`Column1`, `Column2`, …).
- Dos cabeceras con el mismo texto se desambiguan (`Concepto`, `Concepto2`) porque Excel exige
  nombres de columna únicos.

### 4. Lo que se escribe

- `id` correlativo: el mayor `id` de tablas del libro más uno, porque el identificador es único en
  todo el libro.
- `name` y `displayName` con el nombre elegido, `ref` con el rango y `headerRowCount` con 1 o 0.
- Un `<autoFilter>` sobre el rango, como hace Excel por omisión, y el estilo `TableStyleMedium2` con
  bandas de fila.

### 5. Lo que no se modela

- Fila de totales, columnas calculadas, fórmulas de columna y referencias estructuradas.
- Elección del estilo de tabla, que se fija en `TableStyleMedium2`.
- El autofiltro propio de la tabla no se expone como API: se conserva y se genera, pero no se
  modifica.
- No se valida que el contenido de la hoja sea coherente con la tabla (por ejemplo, que la fila de
  cabecera tenga los mismos textos que los nombres de columna después de editarlos).

## Consecuencias

- La API pública nueva es `ExcelTable`, `Worksheet.AddTable`, `Worksheet.Tables` y
  `Worksheet.RemoveTable`.
- El paquete producido puede contener una parte `xl/tables/tableN.xml` y el elemento
  `<tableParts>` de la hoja, con su relación.
- Una tabla ajena se lee con su nombre, su rango, su cabecera y sus columnas, y se conserva al
  guardar.

## Verificación

- `ExcelTableTests`: creación con cabecera, con parte y relación; tabla sin cabecera con nombres
  generados; cabecera con celdas vacías; nombre inválido; nombre repetido; rango inválido; lectura
  de las tablas de un archivo ajeno; preservación al guardar; eliminación de la tabla y de su parte;
  ausencia del diagnóstico de contenido no modelado; libro liberado.
- `OoxmlValidation.CreateXlsxWithTable` construye el fixture con el SDK y `ReadXlsxTables` lee la
  tabla con el SDK.
