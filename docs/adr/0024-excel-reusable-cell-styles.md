# ADR 0024 — Estilos reutilizables de celda

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.4

## Contexto

El incremento 3.3 modeló el formato numérico y las fechas, pero no el resto del formato visual de una
celda. Hasta ahora, cambiar una fuente, un relleno, un borde o una alineación obligaba a manipular el
`styles.xml` a mano o a perder esas propiedades al reescribir la celda.

En SpreadsheetML el formato de una celda es una cadena de referencias:

```
cell/@s ──► cellXfs/xf ──► fonts/font
                       ├─► fills/fill
                       ├─► borders/border
                       └─► alignment (hijo de xf)
```

Varias celdas comparten el mismo `xf` cuando su formato es idéntico. Un estilo reutilizable debe
respetar esa estructura: no basta con escribir propiedades en cada celda, hay que reutilizar los
componentes y el `cellXfs` para no inflar el archivo.

## Decisión

### 1. La API pública es `Cell.Style` y un `CellStyle` inmutable

- `Cell.Style` lee y escribe el formato efectivo de la celda.
- `CellStyle` es un `record` inmutable con propiedades `init` para:
  - fuente: `FontName`, `FontSizePoints`, `Bold`, `Italic`, `Underline`, `FontColor`;
  - relleno: `FillColor` (solo relleno sólido);
  - bordes: `Borders`, un `CellBorders` con las cuatro aristas (`CellBorder?` cada una);
  - alineación: `HorizontalAlignment`, `VerticalAlignment`, `WrapText`.
- Los tipos auxiliares son `CellBorderStyle`, `CellBorder`, `CellBorders`, `CellUnderlineStyle`,
  `CellHorizontalAlignment` y `CellVerticalAlignment`. Sus valores cubren los conjuntos de OOXML que
  esta versión modela (`ST_BorderStyle`, `ST_Underline`, `ST_HorizontalAlignment` y
  `ST_VerticalAlignment`).
- `CellStyle` **no** incluye el formato numérico. `Cell.NumberFormat` sigue siendo la API de 3.3 y se
  conserva al aplicar un estilo. Un mismo `CellStyle` puede aplicarse a celdas con formatos numéricos
  distintos; el `cellXfs` resultante será distinto porque el formato numérico forma parte de él.

### 2. Semántica tri-state y de reemplazo

Cada propiedad de `CellStyle` es tri-state:

- `null` significa «volver al valor por defecto del libro» para esa propiedad.
- `false` en `Bold`, `Italic` o `WrapText` desactiva explícitamente la propiedad.
- Un `CellStyle` vacío (`new CellStyle()`) restablece todas las propiedades modeladas a su valor por
  defecto, pero conserva el formato numérico.

Asignar un estilo **reemplaza** las propiedades modeladas de la celda. No es una operación de
«parche»: `new CellStyle { Bold = true }` deja la fuente sin nombre, tamaño, cursiva, subrayado ni
color explícitos; el relleno sin color; los bordes sin aristas modeladas; y la alineación sin
horizontal, vertical ni ajuste de texto. Lo que no se modela se conserva en la medida descrita en el
apartado 4.

Con una celda ausente:

- el getter devuelve un `CellStyle` vacío y **no crea** la celda;
- asignar un estilo vacío **no crea** la celda;
- asignar un estilo no vacío **crea** la celda para que el formato tenga dónde vivir, aunque la celda
  no tenga valor.

### 3. Creación perezosa y reutilización

- `ExcelWorkbook.Create` sigue sin crear `xl/styles.xml`.
- El primer estilo no vacío, o el primer formato no `General` de 3.3, crea un stylesheet mínimo
  válido con fuente, relleno, borde, `cellStyleXfs`, `cellXfs` y el estilo `Normal`.
- En un libro abierto no se reescribe el stylesheet: se añaden componentes y `cellXfs` al final y se
  actualizan los contadores.
- Cada componente (`font`, `fill`, `border`) y cada `xf` se busca por igualdad exacta de su XML
  (`OuterXml`) antes de añadirlo. Aplicar el mismo estilo a varias celdas con el mismo `xf` base y el
  mismo formato numérico reutiliza un único `cellXfs` y unos únicos componentes.
- Si dos celdas tienen distinto formato numérico, distinto `xf` base o distintas propiedades no
  modeladas, el resultado no es idéntico y se crean entradas distintas. Es la consecuencia correcta de
  conservar el formato efectivo, no un fallo de deduplicación.
- Los componentes y `xf` antiguos no se borran; pueden quedar sin uso, igual que hace Excel.

### 4. Preservación de lo no modelado

Al aplicar un estilo a una celda existente se clona su `xf` actual y solo se cambian las referencias
modeladas. Por tanto se conservan:

- el formato numérico (`numFmtId`, `applyNumberFormat`), `quotePrefix`, `protection`, `extLst` y
  cualquier atributo o hijo no modelado del `xf`;
- las propiedades no modeladas de la fuente: por ejemplo `strike`, `outline`, `shadow`, `condense`,
  `extend`, `vertAlign`, `family`, `charset` y `scheme`. Se eliminan y reescriben solo los hijos
  modelados (`b`, `i`, `u`, `sz`, `color`, `name`);
- las aristas no modeladas del borde, como `diagonal`, `vertical`, `horizontal`, `start` y `end`. Se
  eliminan y reescriben solo `left`, `right`, `top` y `bottom`;
- los atributos no modelados de la alineación, como `indent`, `textRotation`, `shrinkToFit`,
  `readingOrder`, `relativeIndent` y `justifyLastLine`. Se eliminan y reescriben solo `horizontal`,
  `vertical` y `wrapText`;
- el orden de esquema de los hijos de `font` y `border` al insertar los modelados.

El relleno es la excepción: esta versión solo modela relleno sólido. Asignar `FillColor = null`
restablece el relleno a «sin relleno» y asignar un color escribe un `patternFill` sólido; un relleno
con patrón, degradado o color de tema no se conserva como relleno de la celda. Las definiciones
antiguas siguen en el stylesheet, pero la celda deja de referenciarlas.

Tampoco se resuelve la herencia de estilos con nombre (`cellStyleXfs`/`xfId`). El `xf` se conserva y
se lee tal cual; aplicar un estilo vacío a una celda cuyo `xf` 0 hereda de un estilo con nombre puede
no eliminar esa herencia si el `xf` 0 parece vacío.

### 5. Colores

`RgbColor` modela solo el componente RGB. SpreadsheetML escribe colores como ARGB de ocho dígitos:

- al leer se aceptan seis u ocho dígitos; en el segundo caso se ignora el canal alfa;
- al escribir se usa alfa `FF` y se conserva el RGB;
- un color de tema, `auto` o `indexed` sin `rgb` se lee como `null` («sin color RGB explícito»).

### 6. Validación antes de tocar el paquete

`Cell.Style` valida el estilo antes de crear la celda o modificar el stylesheet:

- `FontName`: no vacío, máximo 255 caracteres y caracteres válidos en XML.
- `FontSizePoints`: finito, mayor que 0 y como máximo 409 puntos.
- `Underline`, `HorizontalAlignment` y `VerticalAlignment`: valores definidos del enum.
- `CellBorder` valida su `CellBorderStyle` en el constructor.
- Un estilo inválido lanza `ArgumentException` o `ArgumentOutOfRangeException` y deja el libro sin
  cambios; una celda ausente sigue ausente.
- Un stylesheet existente malformado que impide resolver o añadir el estilo lanza
  `InvalidOfficeFileException`; no se descarta contenido en silencio.

### 7. Límites explícitos

Esta versión no modela ni promete:

- rellenos con patrón, degradados o colores de tema;
- protección de celda, rotación de texto, sangría, `shrinkToFit`, orden de lectura, justificación de
  la última línea ni el resto de atributos de alineación;
- estilos con nombre, estilos de tabla, formato condicional, `dxf`, bordes de tabla o formato de
  número dentro de `CellStyle`;
- aplicar un estilo a un rango de una sola vez (se recorre con `CellRange.Cells` y se asigna
  `Cell.Style` a cada celda);
- conservar el canal alfa o el marcado de tema de un color.

## Consecuencias

- La API pública nueva es `Cell.Style`, `CellStyle`, `CellBorder`, `CellBorders`,
  `CellBorderStyle`, `CellUnderlineStyle`, `CellHorizontalAlignment` y `CellVerticalAlignment`.
- El paquete producido puede incluir nuevas entradas en `fonts`, `fills`, `borders` y `cellXfs`; el
  stylesheet existente no se reescribe.
- Aplicar el mismo estilo a muchas celdas reutiliza componentes y `cellXfs`, que es el objetivo del
  incremento. Aplicar estilos distintos crea entradas distintas.
- El formato numérico y las propiedades no modeladas de la fuente, el borde y la alineación se
  conservan; el relleno no sólido y la herencia de estilos con nombre quedan como limitaciones
  documentadas.

## Verificación

- `ExcelCellStyleTests`: celda ausente sin creación; estilo completo de fuente, relleno, bordes y
  alineación con round-trip por stream; reutilización de `cellXfs` y componentes en varias celdas;
  estilos distintos con `cellXfs` distintos; preservación del formato numérico y de la fecha;
  reseteo con un estilo vacío; creación de una celda sin valor; validación de nombre, tamaño y enums;
  los 14 estilos de borde, los 5 subrayados y las alineaciones horizontales/verticales; lectura de un
  fixture ajeno con fuente, relleno, borde y alineación; preservación de `strike`, `diagonal` e
  `indent` no modelados; libro liberado.
- `ExcelInteropTests.ProducedPackage_WithCellStyles_IsValidAndReadableByTheSdk`: el paquete pasa
  `OpenXmlValidator`; el SDK lee la misma entrada de `cellXfs` para dos celdas y los valores de
  fuente, color, relleno, borde y alineación.
- `OoxmlValidation.CreateXlsxWithFormattedCell` construye los fixtures con el SDK; `ReadXlsxCellStyle`,
  `ReadXlsxCellStyleParts` y `ReadXlsxStyleCounts` leen el resultado con el SDK, no con la librería
  bajo prueba.
