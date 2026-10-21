# ADR 0032 — Hipervínculos y notas

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.9

## Contexto

Dos características de una hoja viven fuera de las celdas:

- Los **hipervínculos** se guardan en `<hyperlinks>` de la hoja. Cada uno cubre un rango (`ref`) y
  apunta a un sitio: o a una dirección externa, con una relación de tipo `hyperlink` y
  `TargetMode="External"`, o a un lugar del propio libro, con el atributo `location`.
- Las **notas** (comentarios) viven en su propia parte (`xl/comments1.xml`), con una lista de autores
  y un `<comment>` por celda, y **necesitan además un dibujo heredado** (`xl/drawings/vmlDrawing1.vml`)
  con la forma de cada nota: Excel usa ese VML para dibujar los indicadores y el cuadro de la nota.

## Decisión

### 1. Hipervínculos

- `Worksheet.Hyperlinks`: catálogo de solo lectura con los hipervínculos de la hoja.
- `Worksheet.AddHyperlink(reference, target, display, tooltip)`: hipervínculo **externo**.
- `Worksheet.AddInternalHyperlink(reference, location, display, tooltip)`: hipervínculo **interno**.
- `Worksheet.RemoveHyperlink(reference)`: elimina el hipervínculo de un rango y su relación.
- `ExcelHyperlink`: expone `Reference`, `Target`, `Location`, `Display`, `Tooltip` e `IsExternal`.

El rango que cubre el hipervínculo se valida como cualquier referencia A1 (una celda o un rango). La
dirección externa tiene que ser una URI absoluta; la API la normaliza con `Uri`, de modo que
`https://example.com` se guarda y se devuelve como `https://example.com/`. Un hipervínculo interno
guarda el texto de la ubicación tal cual, para poder apuntar a `Resumen!A1` o a `A1`.

Al eliminar un hipervínculo externo se elimina también su relación, para no dejar una relación
suelta en la parte de la hoja.

### 2. Notas

- `Worksheet.Comments`: catálogo de solo lectura con las notas de la hoja.
- `Worksheet.AddComment(reference, text, author)`: añade la nota de una celda; si la celda ya tenía
  una, la sustituye.
- `Worksheet.RemoveComment(reference)`: elimina la nota de una celda.
- `ExcelComment`: expone `Reference`, `Text` y `Author`.

La nota se escribe como Excel la escribe: un `run` con la fuente Tahoma de 9 puntos y el color
indexado 81, y el autor se añade a la lista de autores si no estaba. El **dibujo heredado se
reconstruye entero** cada vez que cambia el conjunto de notas, con una forma por nota.

La posición del cuadro de la nota es **aproximada**: se ancla a la celda con los valores que usa
Excel, pero el tamaño y el desplazamiento no se calculan a partir del ancho real de las columnas, que
depende de la fuente por defecto. Excel reposiciona el cuadro al abrir el archivo.

### 3. Lo que no se modela

- Los comentarios con hilo de Excel (los que viven en `xl/threadedComments` y `xl/persons`) y los
  comentarios heredados con formato enriquecido: se conservan, pero no se interpretan.
- El tamaño, la visibilidad y el color del cuadro de la nota.
- Los hipervínculos escritos como fórmula `HYPERLINK()`: son texto, no un hipervínculo del paquete.

## Consecuencias

- La API pública nueva es `ExcelHyperlink`, `ExcelComment`, `Worksheet.Hyperlinks`,
  `Worksheet.AddHyperlink`, `Worksheet.AddInternalHyperlink`, `Worksheet.RemoveHyperlink`,
  `Worksheet.Comments`, `Worksheet.AddComment` y `Worksheet.RemoveComment`.
- El paquete producido puede contener una relación de hipervínculo, la parte de comentarios y la
  parte de dibujo heredado, con sus referencias en la hoja.
- La nota se escribe con el formato de Excel, pero su posición en pantalla es aproximada.

## Verificación

- `ExcelHyperlinkTests`: hipervínculo externo con su relación y su round-trip; hipervínculo interno;
  valores inválidos; eliminación con la relación; lectura de un archivo ajeno; preservación al
  guardar; libro liberado.
- `ExcelCommentTests`: nota con su parte de comentarios y su dibujo heredado; sustitución de la nota
  de una celda; varios autores; valores inválidos; eliminación; lectura de un archivo ajeno;
  preservación al guardar; libro liberado.
- `OoxmlValidation.CreateXlsxWithHyperlinks` y `CreateXlsxWithComments` construyen los fixtures con el
  SDK; `ReadXlsxHyperlinks` y `ReadXlsxComments` los leen con el SDK.
