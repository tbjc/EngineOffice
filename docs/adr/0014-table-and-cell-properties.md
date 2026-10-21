# ADR 0014 — Tablas: propiedades de tabla y de celda

- Estado: Aceptado
- Fecha: 2026-10-18
- Fase: 2, incremento 2.5b

## Contexto

El incremento 2.5 «tablas: crear/editar (filas, celdas, bordes, sombreado, ancho, combinación)» se
dividió en tres con aprobación expresa (2.5a/2.5b/2.5c). El **2.5a** entregó la **estructura y el
contenido** (`docs/adr/0013-tables-structure-and-content.md`); el **2.5c** cubrirá la combinación de
celdas. Este documento fija el **2.5b**: las **propiedades** de la tabla y de la celda.

Hasta 2.5a, el formato de tabla se **preservaba** pero no se modelaba: una tabla ajena con bordes,
sombreado o ancho sobrevivía al editar, pero la API no lo leía ni lo escribía. 2.5b convierte ese
marcado en propiedades de primera clase.

Se verificó con una sonda sobre `DocumentFormat.OpenXml` 3.5.1 que:

- El orden de los hijos de `w:tblPr` es `tblStyle, tblpPr, tblOverlap, bidiVisual, tblW, jc,
  tblCellSpacing, tblInd, tblBorders, shd, tblLayout, tblCellMar, tblLook, tblCaption,
  tblDescription, tblPrChange`, y el de `w:tcPr` es `cnfStyle, tcW, gridSpan, hMerge, vMerge,
  tcBorders, shd, noWrap, tcMar, textDirection, tcFitText, vAlign, hideMark` (+ `tcPrChange`).
- El SDK **no reordena** los hijos: un elemento insertado fuera de orden produce XML inválido.
- La alineación de tabla es `w:tblPr/w:jc`, de tipo `ST_JcTable`, con solo `left`, `center` y `right`
  (el SDK usa `TableRowAlignmentValues`, no `JustificationValues`).
- `w:tblBorders` y `w:tcBorders` solo admiten `top`, `left`, `bottom`, `right`, `insideH`,
  `insideV` (y las diagonales en una celda); el validador del SDK **rechaza** `start`/`end`.
- `w:val` es obligatorio en un borde y en un sombreado: ausente se considera archivo ilegible.
- `w:type` es **opcional** en `w:tblW`/`w:tcW` (ST_TblWidth: «si se omite, se asume `dxa`») y en
  `w:tblLayout` (`fixed`/`autofit`).
- `BorderValues` en el SDK 3.5.1 es un `readonly struct`, **no** un `enum`: no admite
  `Enum.Parse`; se construye desde la cadena que almacena el XML.

## Decisión

### 1. Alcance exacto: seis propiedades de tabla y cuatro de celda

- `Table`: `Width`, `Alignment`, `Layout`, `ShadingColor` y seis bordes (`BorderTop`, `BorderLeft`,
  `BorderBottom`, `BorderRight`, `BorderInsideHorizontal`, `BorderInsideVertical`).
- `TableCell`: `Width`, `VerticalAlignment`, `ShadingColor` y los cuatro bordes exteriores
  (`BorderTop`, `BorderLeft`, `BorderBottom`, `BorderRight`). Una celda **no** expone
  `insideH`/`insideV`: son bordes de la tabla y, aunque el esquema los permita en `w:tcBorders`, no
  tienen sentido en una sola celda. El marcado ajeno de esas aristas y de las diagonales se
  **preserva** sin reescribirse.
- **Fuera de 2.5b**: los márgenes de celda (`w:tblCellMar`/`w:tcMar`) se preservan pero no se
  modelan. Merecen su propio incremento en lugar de colarse en el de propiedades, que cubre lo que
  Word ofrece en su galería de tabla y de celda (ancho, alineación, disposición, bordes, sombreado y
  alineación vertical).

### 2. Tipos de valor en lugar de coordenadas crudas

- `TableWidth` (`readonly struct`): `Type` (`TableWidthType.Auto`/`Points`/`Percentage`) y `Value`
  en puntos o en porcentaje, con fábricas `TableWidth.Points(puntos)`, `.Percentage(pct)` y
  `.Auto`. Las fábricas validan número finito y no negativo y rechazan valores que no caben en el
  entero que OOXML almacena (`ArgumentOutOfRangeException`).
- `TableBorder` (`readonly struct`): `Style` (`BorderStyle`), `SizePoints` y `Color`
  (`RgbColor?`). El ancho se guarda en **octavos de punto** (`w:sz`), la resolución de OOXML: se
  redondea al escribir y se reporta en puntos con esa misma resolución.
- Enumeraciones propias `BorderStyle` (los 27 valores no decorativos de `ST_Border`),
  `TableAlignment`, `TableLayout` y `TableVerticalAlignment`. `BorderStyle` nombra `single` y
  `double` como `SingleLine`/`DoubleLine` para no chocar con tipos de C#, igual que `UnderlineStyle`.
- La API de bordes es **una propiedad por arista** (`table.BorderTop = …`), no un
  `GetBorder(BorderTarget)` con un enumerado de objetivo: cada arista se documenta y se completa con
  IntelliSense por su cuenta, y su nombre dice qué borde es.

### 3. Lectura: tri-state, con error tipado ante lo que no se modela

Toda propiedad es **tri-state**: `null` significa que el elemento **no está** (`w:tblW`, `w:jc`,
`w:tblLayout`, `w:shd`, `w:top`…), que es exactamente como OOXML expresa «lo hereda del estilo o de
los valores predeterminados». Asignar `null` **elimina** el elemento; asignar un valor lo crea si
falta.

La lectura distingue dos fallos, coherente con el resto de la librería:

- **`UnsupportedOfficeFeatureException`** — el archivo usa algo que OOXML define y esta versión no
  modela: la unidad `nil` de `ST_TblWidth`, un patrón de sombreado distinto de `clear`, un estilo de
  borde decorativo (`apples`, `balloons`…). El marcado original **se conserva**.
- **`InvalidOfficeFileException`** — el valor está **mal formado** o falta un atributo obligatorio:
  `w:w`/`w:sz` que no es un número, una medida negativa, un color que no es `#RRGGBB`, un borde o un
  sombreado sin `w:val`, una alineación o una disposición que no es un valor de la enumeración.

Los atributos **opcionales** que falten se leen con su valor por defecto, no como error: `w:tblW`
sin `w:type` es `dxa`, `w:tblLayout` sin `w:type` es `autofit`, un borde sin `w:sz` es medio punto (el
valor que OOXML asume), y un `w:color="auto"` o un `w:fill` ausente o `auto` es «sin color
explícito».

### 4. Escritura: explícita y sin tocar lo que no se modela

- El ancho se escribe siempre con su unidad explícita (`dxa`/`pct`/`auto`), nunca confiando en el
  valor por defecto.
- El sombreado se escribe como relleno sólido con el patrón `clear`.
- El color de un borde sin color explícito se escribe como `auto`.
- Escribir un borde o un sombreado **solo** toca sus atributos (`w:val`, `w:sz`, `w:color`,
  `w:fill`): los atributos no modelados (`w:space`, colores de tema) y las aristas ajenas del mismo
  contenedor quedan intactos.
- El contenedor se crea al primer uso (`w:tblBorders` al fijar el primer borde, `w:tcBorders` al
  fijar el primer borde de la celda) y se **elimina** cuando se queda sin hijos.

### 5. Orden de esquema e inserción

Un borde, un ancho o un sombreado añadido a una tabla o a una celda ajena debe ir **en su sitio**:
el SDK no reordena y la secuencia es obligatoria. Se añadió el ayudante interno `SchemaOrder`
(inserción por tabla de orden) y `BorderEdges` (acceso a las aristas compartidas por `w:tblBorders`
y `w:tcBorders`), y cada propiedad se inserta siguiendo su tabla de orden. Las aristas ajenas
(diagonales) se dejan donde están aunque no estén en la tabla, ordenándose las conocidas a su
alrededor.

`w:tblPr` **no se elimina nunca**, aunque quede vacío: `CT_Tbl` lo **exige** como primer hijo, así que
se crea al primer uso incluso en una tabla mal formada que no lo tuviera. `w:tcPr` es **opcional**: se
crea al primer uso y se elimina cuando la propiedad que se borra lo deja sin hijos ni atributos.

### 6. Un ancho de tabla no cambia la rejilla

`Table.Width` (el ancho **preferido**, `w:tblW`) y la rejilla (`w:tblGrid/w:gridCol`, que fija las
columnas) son cosas distintas. Fijar `Table.Width` escribe solo `w:tblW`; **no** reescribe la
rejilla. Es la misma separación que ya documentaba el 2.5a y la que Word aplica al reconciliar
ambas.

## Consecuencias

- **Nueva capacidad**: `Table.Width`/`Alignment`/`Layout`/`ShadingColor` y sus seis bordes;
  `TableCell.Width`/`VerticalAlignment`/`ShadingColor` y sus cuatro bordes; `TableWidth`,
  `TableBorder`, `BorderStyle`, `TableAlignment`, `TableLayout`, `TableVerticalAlignment`,
  `TableWidthType`.
- Las tablas creadas por `WordDocument.AddTable` siguen naciendo **sin** propiedades (un `w:tblPr`
  vacío y celdas sin `w:tcPr`); solo se escribe lo que el usuario fija.
- Editar estas propiedades **no** altera el resto del marcado de la tabla, de la celda ni del
  documento: sigue siendo válido el criterio del 2.5a de que las tablas existentes no se destruyen.
- Los huecos quedan anotados en `docs/compatibility-matrix.md` y en la nota de alcance de
  `docs/roadmap.md`. En particular, los **márgenes de celda** siguen sin modelarse.
- Alternativas descartadas: un único `GetBorder`/`SetBorder` con un enumerado `BorderTarget`
  (opaco en el punto de uso y sin poder documentar cada arista); tratar `pct` o `auto` como
  «no soportado» (son válidos y triviales de modelar); tratar un `w:type` omitido como
  «no soportado» (el esquema define el valor por defecto, y rechazarlo volvería ilegible un archivo
  válido); modelar `start`/`end` como aristas (el SDK las rechaza en este contexto); escribir las
  aristas `insideH`/`insideV` de una celda (no son suyas).
