# ADR 0013 — Tablas: estructura y contenido

- Estado: Aceptado
- Fecha: 2026-10-17
- Fase: 2, incremento 2.5a

## Contexto

El incremento 2.5 «tablas: crear/editar (filas, celdas, bordes, sombreado, ancho, combinación)» es
grande. Con aprobación expresa del interesado se dividió en tres incrementos, igual que se hizo con
el 2.2/2.2b, **en lugar de declarar hecho algo a medias**:

- **2.5a (este documento)** — estructura y contenido: crear tablas, y leer y editar filas, celdas y
  su texto, sin destruir el resto del marcado.
- **2.5b** — propiedades: ancho (de tabla y de celda), alineación, disposición (`w:tblLayout`),
  bordes, sombreado y alineación vertical.
- **2.5c** — combinación de celdas (`w:gridSpan`/`w:vMerge`).

Hasta 2.5a las tablas eran **contenido de cuerpo no modelado**: se preservaban y se reportaban con
`WORD-0001-UNMODELLED-BODY-CONTENT`. `WordDocument.Paragraphs` y `Text` solo veían los `w:p`
directos del cuerpo.

Se verificó con una sonda sobre `DocumentFormat.OpenXml` 3.5.1 que:

- `w:tbl` tiene los hijos, en orden, `w:tblPr`, `w:tblGrid` y `w:tr*`. **`w:tblPr` y `w:tblGrid` son
  obligatorios** (el validador da error si faltan); un `w:tblPr` vacío y un `w:gridCol` sin `w:w`
  validan con **0 errores**, igual que una `w:tc` sin `w:p` y una tabla sin filas o sin columnas.
- `w:tc` contiene `w:tcPr?` y contenido de bloque (`w:p` y `w:tbl`). Una tabla anidada dentro de una
  celda valida.
- El orden de `w:tcPr` es `cnfStyle, tcW, gridSpan, hMerge, vMerge, tcBorders, shd, noWrap, tcMar,
  textDirection, tcFitText, vAlign, hideMark`; el orden de `w:tblPr` es `tblStyle, tblpPr,
  tblOverlap, bidiVisual, tblW, jc, tblCellSpacing, tblInd, tblBorders, shd, tblLayout, tblCellMar,
  tblLook, tblCaption, tblDescription, tblPrChange`.
- **El SDK no reordena** los hijos de `w:tcPr`/`w:tblPr`: construir con las propiedades en orden
  incorrecto produce XML inválido. Hace falta insertar en orden de esquema (como ya se hace con
  `w:pPr`/`w:rPr`).
- `w:tblBorders`/`w:tcBorders` usan `top/left/bottom/right/insideH/insideV`; `w:tblCellMar` acepta
  `w:left`/`w:right` (el validador **rechaza** `w:start`/`w:end`).
- Un cuerpo que termina en tabla valida.

## Decisión

### 1. La tabla es un elemento de bloque al mismo nivel que el párrafo

`w:body` intercala `w:p` y `w:tbl`. Se añaden `WordDocument.Tables` (tablas **directas** del cuerpo,
en orden de documento) y `WordDocument.AddTable(filas, columnas)`, que **añade al final** del cuerpo,
igual que `AddParagraph`. No se expone una lista unificada de bloques: `Paragraphs` y `Tables`
conservan el contrato «solo hijos directos del cuerpo», y `AddTable`/`AddParagraph` solo añaden al
final. Insertar o mover bloques sigue fuera de alcance.

`AddTable` exige `filas >= 1` y `columnas >= 1` (`ArgumentOutOfRangeException` en caso contrario).
Crea, en orden de esquema:

```xml
<w:tbl>
  <w:tblPr/>
  <w:tblGrid><w:gridCol w:w="2400"/>…</w:tblGrid>
  <w:tr>…<w:tc><w:p/></w:tc>…</w:tr>
</w:tbl>
```

`w:tblPr` se deja **vacío** (no se escribe `w:tblW` ni bordes) porque las propiedades no se modelan
todavía; es válido y deja el comportamiento de la tabla en manos de los valores predeterminados de
Word. Cada `w:gridCol` usa un ancho fijo predeterminado (`DefaultColumnWidthTwips` = 2400 twips,
≈ 4,2 cm); **no** se mide la página, porque el ancho no se modela todavía. Cada celda empieza con un
párrafo vacío, que es lo que Word necesita para editar la celda.

### 2. El modelo es una vista ligera de solo los hijos directos

- `Table`: `Rows`, `RowCount`, `ColumnCount` (número de `w:tblGrid/w:gridCol`; `0` si falta la
  rejilla, que es lo único que puede faltar en una tabla mal formada), `AddRow`.
- `TableRow`: `Cells`, `CellCount`.
- `TableCell`: `Paragraphs` (solo los `w:p` directos), `Text` (lectura y escritura), `AddParagraph`.

`Table.AddRow` añade una fila con `ColumnCount` celdas vacías al final de las filas existentes.
Las **columnas se fijan al crear la tabla**: no se añaden ni se quitan columnas, porque eso obligaría
a mantener la rejilla y es una edición estructural mayor.

### 3. Escritura de texto que nunca descarta contenido no modelado

`TableCell.Text` (lectura) concatena el texto de sus párrafos separándolos por `\n`. El *setter*
reemplaza **solo los párrafos** de la celda por uno nuevo con el texto; **conserva** `w:tcPr` y
cualquier otro contenido de bloque (por ejemplo, una tabla anidada), a diferencia del *setter* de
`Paragraph.Text`, que sí elimina el contenido que llevaba el texto antiguo. La razón es la regla de
integridad de datos: asignar texto a una celda no debe borrar en silencio una tabla anidada que la
versión no modela. El texto se valida con `PlainText.EnsureValid` (nada de tabuladores, saltos ni
caracteres que OOXML no puede almacenar).

### 4. `WordDocument.Text` incluye el texto de las celdas

Antes de 2.5a `Text` unía, con `\n`, el texto de cada párrafo de nivel superior; el texto de una
tabla se perdía en silencio (solo lo advertía el diagnóstico de contenido no modelado). Con las
tablas modeladas, `Text` recorre el cuerpo en orden y une con `\n` los **bloques**: cada párrafo de
nivel superior aporta su texto, y cada tabla aporta una **línea por fila** con las celdas separadas
por **tabulador** (una celda con varios párrafos aporta varias líneas). El contenido que no se
modela (una `w:sdt`, una tabla anidada) no aporta texto; se preserva y se advierte por diagnóstico.

### 5. El contenido no modelado dentro de una tabla se preserva

Editar el cuerpo, añadir filas o reescribir el texto de una celda no altera el resto del marcado de
la tabla ni de sus celdas: bordes, sombreado, anchos, `w:gridSpan`, `w:vMerge` y toda propiedad que
2.5b/2.5c modelarán más adelante sobreviven. Es la materialización del criterio de aceptación del
incremento: «tablas existentes no se destruyen al editar».

## Consecuencias

- **Cambio de contrato en `WordDocument.Text`**: pasa a incluir el texto de las celdas. Es una
  ampliación deliberada y documentada; las pruebas y el `README` se actualizan con ella. Ninguna otra
  parte de la API cambia: `Paragraphs` sigue devolviendo solo párrafos de nivel superior.
- **Nueva capacidad**: `WordDocument.Tables`/`AddTable`, `Table`, `TableRow`, `TableCell`.
- `WORD-0001-UNMODELLED-BODY-CONTENT` deja de aplicarse a `w:tbl`; sigue aplicándose a otros
  elementos de bloque (`w:sdt`, `w:sectPr`, …).
- Los huecos quedan anotados en `docs/compatibility-matrix.md` y en la nota de alcance de
  `docs/roadmap.md`.
- Alternativas descartadas: modelar una lista unificada de bloques del cuerpo (mayor cambio de API,
  no necesario para el criterio); escribir un `w:tblW` predeterminado (propiedad no modelada, se
  deja para 2.5b); hacer que `TableCell.Text` borre todo salvo `w:tcPr` (pérdida silenciosa de
  tablas anidadas).
