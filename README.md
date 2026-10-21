# OfficeEngine

Librería .NET independiente para **crear, leer, editar y guardar** documentos Word (`.docx`) y
libros Excel (`.xlsx`), y para **convertirlos a PDF** con un motor externo detrás de una interfaz.

> **Estado: fases 0 a 6 completadas; la fase 7 (publicación) está pendiente.** La Fase 2 cubre Word
> (formato de carácter y de párrafo, tabulaciones, saltos, estilos —referencias y catálogo de solo
> lectura—, listas numeradas y con viñetas, tablas —estructura, contenido, propiedades y combinación
> de celdas—, secciones —tamaño/orientación y márgenes—, encabezados y pies por sección, imágenes
> inline incrustadas y relaciones, hipervínculos externos, propiedades básicas del documento,
> búsqueda y reemplazo, y campos con numeración de páginas). La Fase 3 cubre Excel (acceso a
> celdas por índice, rangos, fechas, formatos numéricos, estilos reutilizables, escritura de
> fórmulas, rangos con nombre, validación de datos, autofiltros, tablas estructuradas, paneles
> congelados, configuración de impresión, hipervínculos, notas, imágenes, gráficos y tabla de
> cadenas compartidas). La Fase 4 añade la conversión a PDF, la Fase 5 la carga de OOXML hostil, el
> corpus de corruptos, las mediciones y la regresión, y la Fase 6 la conversión de los formatos
> heredados y el motor de fórmulas. Todo lo que se declara soportado está respaldado por
> pruebas automatizadas. Consulta [`docs/compatibility-matrix.md`](docs/compatibility-matrix.md)
> para el estado real, capacidad por capacidad. No se afirma paridad con otras librerías (por
> ejemplo Spire) ni con Microsoft Office.
>
> La documentación de uso está en el sitio de Docusaurus de [`docs-site/`](docs-site): índice,
> Word, Excel, PDF y formatos heredados, con un ejemplo por función.

## Qué funciona hoy

**Word**

- Crear un documento, añadir párrafos con texto, guardarlo y reabrirlo (por archivo y por stream).
- Leer y editar el texto de cada párrafo.
- Modelar cada párrafo como una secuencia de **runs** (`Paragraph.Runs`, `Paragraph.AddRun`) y dar
  a cada run **formato de carácter**: familia y tamaño de fuente, negrita, cursiva, subrayado (los
  18 valores de OOXML) y color RGB. Las propiedades que admiten herencia son tri-state
  (heredar / desactivado explícito / activado).
- Dar **formato de párrafo**: alineación (izquierda, centrada, derecha, justificada y «both»),
  sangría (primera línea, izquierda, derecha y francesa) y espaciado (anterior, posterior e
  interlineado exacto), todo en puntos. Las propiedades se insertan en el orden que exige el
  esquema de `w:pPr`.
- Definir **tabulaciones** de párrafo (`Paragraph.TabStops`, `AddTabStop`, `ClearTabStops`) con
  posición en puntos, alineación (izquierda, centrada, derecha, decimal, barra y «clear») y relleno
  (puntos, guiones, subrayado…), y componer el contenido de un run con
  `Run.AddTab`/`Run.AddBreak` (salto de línea, de página o de columna). Los tabuladores y saltos no
  son texto: escribirlos como carácter en `Run.Text`/`Paragraph.Text` se rechaza con error tipado.
- Referenciar **estilos** desde un párrafo (`Paragraph.StyleId`, `w:pStyle`) o un run
  (`Run.StyleId`, `w:rStyle`) y **leer el catálogo** del documento (`WordDocument.Styles`,
  `WordDocument.FindStyle`) con el formato que aporta cada definición (`OwnFormatting`) y el
  resuelto por la herencia `w:basedOn` (`EffectiveFormatting`). Es **solo lectura**: no se crean ni
  editan definiciones y `styles.xml` se conserva tal cual.
- Crear **listas** numeradas y con viñetas (`WordDocument.AddList`), situar cada párrafo en una lista
  y en un nivel (`Paragraph.ListId`/`ListLevel`, `Paragraph.ApplyList`) y **leer el catálogo** del
  documento (`WordDocument.Lists`, `WordDocument.FindList`) con el formato, el texto y la justificación
  de cada nivel. Un documento nuevo sin listas no inventa `numbering.xml`, y el existente se conserva
  tal cual salvo las listas que se crean.
- Crear **tablas** (`WordDocument.AddTable(filas, columnas)`) y recorrer y editar su **estructura y su
  contenido**: `Table.Rows`/`RowCount`/`ColumnCount`, `TableRow.Cells`, `TableCell.Paragraphs`,
  `TableCell.Text` y `Table.AddRow`. Una tabla es un elemento de bloque al mismo nivel que un párrafo,
  y el texto de las celdas forma parte de `WordDocument.Text`.
- Dar **formato a la tabla y a la celda**: ancho preferido (`TableWidth`, en puntos, porcentaje o
  `auto`), alineación y disposición de la tabla (`TableAlignment`, `TableLayout`), alineación
  vertical de la celda, sombreado (`ShadingColor`) y bordes (`TableBorder` con estilo, grosor en
  puntos y color), arista por arista. Cada propiedad es tri-state: `null` significa «no se declara»
  y asignarlo la elimina. Un valor que OOXML define pero la versión no modela (unidad `nil`, patrón
  de sombreado no sólido, borde decorativo) **falla con error tipado** en vez de inventarse, y el
  marcado original se conserva.
- **Combinar y dividir celdas**: leer y escribir las marcas (`TableCell.ColumnSpan` para
  `w:gridSpan` y `TableCell.VerticalMerge` para `w:vMerge`) y operar con
  `TableRow.MergeCells`/`UnmergeCell` (horizontal) y
  `Table.MergeCellsVertically`/`UnmergeCellsVertically` (vertical). Combinar en horizontal **mueve**
  el contenido de las celdas que absorben, así que no se pierde ni un párrafo ni una tabla anidada;
  combinar en vertical solo escribe las marcas. Las combinaciones **rectangulares** se rechazan con
  error tipado en vez de producir una tabla rota.
- **Secciones**: leer todas las secciones (`WordDocument.Sections`), añadir una al final
  (`WordDocument.AddSection`) y dar a cada una su **tamaño y orientación de página**
  (`Section.PageSize`, `PageSize`, con `Rotate()` para pasar de vertical a horizontal), sus
  **márgenes** (`Section.Margins`, `PageMargins`) y su **tipo de inicio** (`Section.StartType`). Un
  documento nuevo empieza con una sección US Letter vertical y márgenes de una pulgada; el marcado de
  sección que no se modela (columnas, `w:docGrid`…) se conserva.
- **Encabezados y pies por sección**: leer la parte de cada tipo (`Section.Header`/`Footer`,
  `HeaderFooterType`), crearla con `AddHeader`/`AddFooter` y quitarla con `RemoveHeader`/`RemoveFooter`
  sin borrar la parte, con contenido editable (párrafos, tablas y texto). Varias secciones pueden
  **compartir** la misma parte (vínculo sección-parte) y la primera página se distingue con
  `Section.DifferentFirstPage` (`w:titlePg`).
- **Imágenes inline incrustadas y relaciones**: añadirlas desde archivo o stream
  (`Paragraph.AddImage`, `ImageFormat`), con el tamaño en puntos convertido a EMU, y leerlas
  (`Paragraph.Images`, `Image`) con su tamaño, tipo de contenido y bytes. Funciona en el cuerpo, en
  celdas de tabla y en encabezados/pies; cada imagen crea su propia parte y relación, y las
  relaciones no modeladas no se rompen al editar.
- **Hipervínculos externos**: crearlos (`Paragraph.AddHyperlink`), leerlos (`Paragraph.Hyperlinks`,
  `Hyperlink`) y editar su texto (`Hyperlink.Text`, `Hyperlink.Runs`) y su destino
  (`Hyperlink.Destination`) en el cuerpo, celdas de tabla y encabezados/pies. Editar un destino crea
  una relación nueva para ese enlace, respeta las relaciones compartidas y elimina la anterior cuando
  queda sin uso. Los enlaces internos a marcadores (`w:anchor`) se preservan, pero no se modelan.
- **Propiedades básicas del documento**: leer y editar título, asunto, autor, palabras clave,
  descripción, último editor, revisión, fechas de creación/modificación/impresión, categoría, estado,
  identificador, idioma y versión con `WordDocument.Properties` (`DocumentProperties`). Se guardan en
  `docProps/core.xml`, que se crea solo al asignar el primer valor. Las propiedades extendidas
  (`docProps/app.xml`) y personalizadas (`docProps/custom.xml`) se preservan, pero no se modelan.
- **Búsqueda y reemplazo** con alcance explícito: `WordDocument.FindText`/`ReplaceText` y
  `Paragraph.FindText`/`ReplaceText`. Buscan en el cuerpo, sin cruzar párrafos, y el texto
  reemplazado conserva el formato del primer run de la coincidencia, aunque el texto esté repartido
  en varios runs.
- **Campos y numeración de páginas**: escribir un campo en cualquier párrafo
  (`Paragraph.AddField("PAGE")`) o en el pie (`HeaderFooter.AddPageNumber`). Los campos se **escriben,
  no se evalúan**: el resultado guardado es el que se da y Word lo recalcula al abrir o imprimir.
- El contenido que la API no modela (definiciones de estilo y de lista, márgenes de celda,
  `w:hMerge` heredado, tablas anidadas, imágenes flotantes o vinculadas, gráficos, enlaces internos a
  marcadores, otros atributos de `w:rPr`, propiedades extendidas/personalizadas… etc.) se **preserva** al guardar y, cuando aplica, se reporta por
  diagnósticos: no se descarta en silencio.

**Excel**

- Crear un libro, añadir/eliminar/renombrar hojas.
- Leer y escribir celdas por referencia A1 o por **índice de fila/columna (1-based)**, con texto,
  número, booleano, fecha/hora o vacío; cada `Cell` expone su `Row`, su `Column` y su `Reference`.
- Leer y escribir **rangos** (`A1:B2`) por referencia o por índices, recorrer sus celdas en orden de
  fila y leer o escribir una matriz de valores con `GetValues`/`SetValues`.
- Escribir **fechas y horas** como serial de Excel con un formato explícito (`Cell.SetDateTime`) y
  leerlas como `DateTime` cuando la celda tiene un formato de fecha/hora.
- Leer y escribir **formatos numéricos** integrados y personalizados con `Cell.NumberFormat`,
  preservando los estilos que esta versión no modela.
- Leer y escribir **estilos reutilizables** de celda (`Cell.Style`, `CellStyle`): fuente, relleno
  sólido, bordes por arista y alineación. Aplicar el mismo estilo a varias celdas reutiliza
  `fonts`/`fills`/`borders`/`cellXfs`; el formato numérico y las propiedades no modeladas de fuente,
  borde y alineación se conservan.
- Leer valores de celdas con cadenas compartidas y el valor en caché de una fórmula.
- Escribir **fórmulas** con `Cell.Formula` (se guardan como fórmulas, sin calcularlas): el resultado
  almacenado se conserva al leer y se declara ausente al escribir, y el libro pide un recálculo al
  abrirse. Escribir un valor literal sobre una celda con fórmula elimina la fórmula.
- Crear y leer **rangos con nombre** (`ExcelWorkbook.DefinedNames`, `Worksheet.AddNamedRange`): con
  ámbito de libro o de hoja. Los nombres que no son rangos (constantes, fórmulas, áreas de
  impresión) se listan y se conservan.
- Crear y leer **reglas de validación de datos** (`Worksheet.AddDataValidation`): lista de valores,
  número entero o decimal con límites. Las reglas de otros libros se leen y se conservan.
- Escribir el **autofiltro** de una hoja (`Worksheet.AutoFilter`) y leer el de un libro ajeno.
- Crear y leer **tablas estructuradas** (`Worksheet.AddTable`, `Worksheet.Tables`): con parte propia,
  cabecera y nombres de columna.
- **Congelar paneles** (`Worksheet.FreezePanes`), **ocultar filas y columnas** y fijar su **alto y
  ancho** (`RowHeight`, `ColumnWidth`, `IsRowHidden`, `IsColumnHidden`).
- Describir la **impresión** de la hoja: márgenes, orientación, tamaño de papel, escala, ajuste a
  páginas y opciones de impresión (`Worksheet.PageMargins`, `PageSetup`, `PrintOptions`).
- Escribir **hipervínculos** externos e internos (`Worksheet.AddHyperlink`, `AddInternalHyperlink`) y
  **notas** en las celdas (`Worksheet.AddComment`), con su relación y su parte de comentarios.
- Anclar **imágenes** en una celda (`Worksheet.AddImage`), en PNG, JPEG, GIF, BMP o TIFF, con su
  parte de medios y su dibujo.
- Escribir **gráficos** de columna, barra, línea o sector (`Worksheet.AddChart`), con una serie y su
  caché de datos.
- **Calcular fórmulas** con el motor propio (`OfficeEngine.Formulas`, 34 funciones en inglés y español).
  `ExcelWorkbook.EvaluateFormula("=A1*B1")` evalúa sobre los **valores** del libro; el motor no sigue
  dependencias ni recalcula nada. La matriz de lo que se soporta está en
  [`docs/formula-matrix.md`](docs/formula-matrix.md).
- **Convertir los formatos heredados** de Word (`.doc`) y Excel (`.xls`) a sus equivalentes modernos
  con `LegacyConverter` (proyecto `OfficeEngine.Legacy`). Es una **conversión**, no una lectura: el
  resultado se abre con la API moderna, y las macros, los objetos incrustados y la maquetación
  exacta del original no se conservan.
- Escribir el texto en la **tabla de cadenas compartidas**, de modo que el texto repetido se guarda
  una sola vez.
- Preservar el contenido no modelado (por ejemplo `mergeCells`) y reportarlo por diagnósticos.

**Transversal**

- Límites de recursos frente a entradas no confiables (tamaño, número de entradas, bomba ZIP,
  **profundidad de XML** y rechazo de DTD y entidades externas). Un paquete con XML anidado en exceso
  falla con un error tipado, en lugar de terminar el proceso.
- Errores tipados, diagnósticos y guardado atómico a archivo.
- **Rendimiento medido**: el proyecto `OfficeEngine.Benchmarks` mide la creación, la apertura y el
  round-trip de Word y Excel, y lo compara con una referencia (`benchmarks/baseline.txt`).

**PDF**

- **Convertir a PDF** con un motor externo, a través de la interfaz `IPdfConverter` (proyecto
  `OfficeEngine.Pdf`) y del adaptador `LibreOfficePdfConverter` (proyecto
  `OfficeEngine.Pdf.LibreOffice`), que usa **LibreOffice en modo headless**.
- La conversión parte de los **bytes que se le dan**: para exportar el estado actual de un
  documento, se guarda en memoria y se entrega ese flujo. El motor corre aislado, con su propio
  perfil y sus propios temporales, y el resultado se comprueba (código de salida y cabecera `%PDF-`).
- **No hay equivalencia visual con Microsoft Office**: el PDF depende de la versión de LibreOffice,
  de las fuentes instaladas y de la configuración del sistema.

## Lo que **no** funciona todavía

Consulta la matriz para el detalle. En resumen: crear/editar/eliminar definiciones de estilo y
personalizar niveles de lista (la API crea listas de un solo nivel), las **combinaciones de celdas
rectangulares**, los **márgenes de celda**, el `w:hMerge` heredado y las tablas anidadas, las
**columnas de sección** y el **borrado/movimiento de secciones**, el contenido no modelado de los
encabezados/pies (controles de contenido y otras estructuras), `w:evenAndOddHeaders` (las páginas
pares), las imágenes flotantes o vinculadas, los gráficos, los hipervínculos internos a marcadores,
las propiedades extendidas o personalizadas del documento; la búsqueda en encabezados, pies, notas y
comentarios, y las expresiones regulares; resaltado, tachado y colores de tema de carácter; los valores que la versión no modela y por
tanto **fallan con un error tipado en
lugar de inventarse** (alineaciones `start`/`end`/kashida, interlineado múltiple en
`LineSpacingPoints`, sangrías en caracteres, formatos de numeración que no se modelan, unidades y
patrones de tabla que no se modelan); rellenos con patrón, degradados o color de tema, estilos con
nombre, estilos de tabla y formato condicional en Excel; los criterios de filtrado, la fila de
totales y las columnas calculadas de las tablas; las
fórmulas compartidas, de matriz y de tabla (se conservan, pero no se escriben ni se expanden); los
formatos integrados de fecha dependientes de la configuración regional no se
exponen como código; **cálculo de fórmulas y de campos**; las **opciones del PDF** (páginas, marcas
de agua, cifrado, metadatos); **DOC/XLS**; y **documentos
cifrados**.

## Requisitos

- .NET 8 (SDK 8.0.400 o superior).
- Para convertir a PDF: **LibreOffice** instalado (se comprueba al usar el adaptador; las pruebas
  de integración se omiten con un motivo explícito si no está).

## Uso

```csharp
using System;
using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Word;

// Word: crear, escribir y guardar.
using (WordDocument document = WordDocument.Create())
{
    document.AddParagraph("Informe mensual");

    // Un párrafo puede tener varios runs, cada uno con su propio formato de carácter.
    Paragraph summary = document.AddParagraph(string.Empty);
    Run amount = summary.AddRun("Cifras al alza");
    amount.Bold = true;
    amount.Underline = UnderlineStyle.SingleLine;
    amount.FontSizePoints = 11.0;
    amount.Color = new RgbColor(0x1F, 0x49, 0x7D);
    summary.AddRun(" frente al mes anterior.");

    // El formato de párrafo se aplica al párrafo completo (alineación, sangría y espaciado).
    summary.Alignment = ParagraphAlignment.Center;
    summary.LeftIndentPoints = 18.0;
    summary.SpacingAfterPoints = 6.0;

    // Las tabulaciones pertenecen al párrafo; se invocan con Run.AddTab. Un salto (de línea,
    // de página o de columna) se añade con Run.AddBreak. No se escriben como texto.
    Paragraph prices = document.AddParagraph("Precio");
    prices.AddTabStop(new TabStop(72.0, TabStopAlignment.Decimal, TabStopLeader.Dot));
    prices.AddRun(string.Empty).AddTab().AddText("1.234,56");

    // Un párrafo puede nombrar un estilo de párrafo y un run uno de carácter (w:pStyle/w:rStyle).
    summary.StyleId = "Heading1";
    amount.StyleId = "Emphasis";

    // La sección define el tamaño/orientación de página y los márgenes. Un documento nuevo ya
    // tiene una sección US Letter; aquí se cambia a A4. `AddSection` añade otra al final.
    document.Sections[0].PageSize = PageSize.A4;
    document.Sections[0].Margins = new PageMargins(72.0, 72.0, 72.0, 72.0);

    // Un encabezado o pie es una parte de la sección; `AddHeader` la crea por tipo y nace con un
    // párrafo vacío. Varias secciones pueden compartir la misma parte.
    document.Sections[0].AddHeader(HeaderFooterType.Primary).Paragraphs[0].Text = "Informe mensual";

    // El pie lleva el número de página: el campo se escribe, no se calcula. Word lo recalcula al
    // abrir o al imprimir; el resultado guardado es el que se le da.
    HeaderFooter footer = document.Sections[0].AddFooter(HeaderFooterType.Primary);
    footer.Paragraphs[0].AddRun("Página ");
    footer.Paragraphs[0].AddField("PAGE", "1");

    // Se crea una lista (numerada o con viñetas) con `document.AddList` y se aplica a un párrafo con
    // `Paragraph.ApplyList`; el nivel es opcional (0..8).
    WordList items = document.AddList(WordListFormat.Bullet);
    document.AddParagraph("Primer punto").ApplyList(items);
    document.AddParagraph("Segundo punto").ApplyList(items);

    // Se crea una tabla con `document.AddTable(filas, columnas)`; cada celda empieza con un párrafo
    // vacío que se rellena con `TableCell.Text`. La tabla se añade al final del cuerpo, al mismo
    // nivel que los párrafos.
    Table totals = document.AddTable(2, 2);
    totals.Rows[0].Cells[0].Text = "Concepto";
    totals.Rows[0].Cells[1].Text = "Importe";
    totals.Rows[1].Cells[0].Text = "Total";
    totals.Rows[1].Cells[1].Text = "1.234,56";

    // Combinar celdas: `MergeCells` une el rango en la primera celda (con su contenido) y
    // `MergeCellsVertically` combina un tramo de filas en una columna. `UnmergeCell` las separa.
    TableCell header = totals.Rows[0].Cells[0];
    totals.Rows[0].MergeCells(0, 2);
    header.Text = "Concepto e importe";
    Console.WriteLine(header.ColumnSpan); // 2

    document.Save("informe.docx");
}

// Word: abrir, editar y guardar en otro archivo.
using (WordDocument loaded = WordDocument.Open("informe.docx"))
{
    loaded.Paragraphs[0].Text = "Informe final";

    // Leer el formato de carácter de un run.
    Run amount = loaded.Paragraphs[1].Runs[0];
    Console.WriteLine(amount.Bold);            // True
    Console.WriteLine(amount.FontSizePoints);  // 11
    Console.WriteLine(amount.Color);           // #1F497D

    // Leer el formato de párrafo (se aplica al párrafo completo).
    Console.WriteLine(loaded.Paragraphs[1].Alignment);          // Center
    Console.WriteLine(loaded.Paragraphs[1].LeftIndentPoints);   // 18
    Console.WriteLine(loaded.Paragraphs[1].SpacingAfterPoints); // 6

    // Buscar y reemplazar en el cuerpo. El texto reemplazado conserva el formato del primer run
    // de la coincidencia, aunque el texto esté repartido en varios runs.
    loaded.ReplaceText("borrador", "final");
    foreach (WordTextMatch match in loaded.FindText("final"))
    {
        Console.WriteLine($"{match.Paragraph} @ {match.StartIndex}");
    }

    // Leer las tabulaciones de un párrafo.
    TabStop stop = loaded.Paragraphs[2].TabStops[0];
    Console.WriteLine($"{stop.PositionPoints} pt {stop.Alignment} {stop.Leader}"); // 72 pt Decimal Dot

    // Leer las referencias de estilo (el estilo puede estar definido en el documento o no).
    Console.WriteLine(loaded.Paragraphs[1].StyleId);         // Heading1
    Console.WriteLine(loaded.Paragraphs[1].Runs[0].StyleId); // Emphasis

    // Leer la sección: tamaño/orientación y márgenes, en puntos.
    Console.WriteLine(loaded.Sections[0].PageSize);  // 595.3 × 841.9 pt (Portrait)
    Console.WriteLine(loaded.Sections[0].Margins);   // top 72 pt, right 72 pt, bottom 72 pt, left 72 pt

    // Leer el encabezado y el pie de una sección (null si la sección no declara ese tipo).
    Console.WriteLine(loaded.Sections[0].Header(HeaderFooterType.Primary)!.Text); // Informe mensual
    Console.WriteLine(loaded.Sections[0].Footer(HeaderFooterType.Primary)!.Text); // Página

    // Leer la lista de un párrafo y el formato del nivel al que pertenece.
    Console.WriteLine(loaded.Paragraphs[3].ListId);   // 1
    Console.WriteLine(loaded.Paragraphs[3].ListLevel); // 0
    Console.WriteLine(loaded.FindList(1)!.Levels[0].Format); // Bullet

    // Leer una tabla: filas, celdas y su texto. `AddRow` añade una fila con una celda por columna.
    Table totals = loaded.Tables[0];
    Console.WriteLine(totals.RowCount);                 // 2
    Console.WriteLine(totals.Rows[0].Cells[0].Text);     // Concepto
    TableRow note = totals.AddRow();
    note.Cells[0].Text = "Nota";
    note.Cells[1].Text = "0";

    loaded.Save("salida.docx");

    foreach (var diagnostic in loaded.Diagnostics)
    {
        Console.WriteLine(diagnostic); // p. ej. contenido preservado pero no modelado
    }
}

// El catálogo de estilos de un documento con `word/styles.xml` se lee con `document.Styles` y
// `document.FindStyle(id)`. Cada `WordStyle` expone `Id`, `Name`, `Type`, `IsDefault`, `BasedOnId`
// y el formato propio (`OwnFormatting`) o el resuelto por la herencia `w:basedOn`
// (`EffectiveFormatting`). Es **solo lectura**: no se crean ni editan definiciones.

// El catálogo de listas de un documento con `word/numbering.xml` se lee con `document.Lists` y
// `document.FindList(id)`. Cada `WordList` expone `Id`, `AbstractNumberId` y sus `Levels` (con el
// formato, el texto y la justificación de cada nivel). `document.AddList(formato)` crea una lista
// nueva de un solo nivel; `Paragraph.ListId`/`ListLevel` (w:numPr) sitúan el párrafo en una lista.

// Las tablas se crean con `document.AddTable(filas, columnas)` y se recorren con `Table.Rows`,
// `TableRow.Cells` y `TableCell.Paragraphs`. `TableCell.Text` lee y escribe el texto de una celda y
// `Table.AddRow` añade una fila con una celda por columna. El texto de las celdas forma parte de
// `document.Text`. El formato se da con `Table.Width`/`Alignment`/`Layout`/`ShadingColor`/`BorderX`
// y `TableCell.Width`/`VerticalAlignment`/`ShadingColor`/`BorderX`. La combinación de celdas se lee
// con `TableCell.ColumnSpan`/`VerticalMerge` y se edita con `TableRow.MergeCells`/`UnmergeCell` y
// `Table.MergeCellsVertically`/`UnmergeCellsVertically`. Lo que no se modela (márgenes de celda,
// `w:hMerge` heredado) y las tablas anidadas se preservan.

// Excel: crear, escribir celdas y guardar.
using (ExcelWorkbook workbook = ExcelWorkbook.Create())
{
    Worksheet sheet = workbook.Worksheets.Add("Resumen");

    // Un rango incluye las dos esquinas y se puede escribir como matriz.
    sheet.GetRange("A1:B2").SetValues(new object?[,]
    {
        { "Concepto", "Importe" },
        { "Total", 270.5 },
    });

    // Una celda también se direcciona por índice 1-based de fila y columna: [3, 1] es A3.
    sheet[3, 1].Value = "IVA";
    sheet[3, 2].Value = 56.8;
    sheet[3, 2].NumberFormat = "0.00";

    // Una fecha se guarda como serial y necesita un formato explícito.
    sheet[4, 1].SetDateTime(new DateTime(2024, 1, 31), "yyyy-mm-dd");

    // Una fórmula se guarda como fórmula y no se calcula: el motor que abra el archivo la recalcula.
    sheet["C1"].Formula = "=B1*1.16";

    // Un rango con nombre se puede leer por su nombre; el ámbito puede ser el libro o la hoja.
    sheet.AddNamedRange("Tabla", "A1:B2");

    // Una validación de datos, un autofiltro y una tabla estructurada.
    sheet.AddDataValidation(ExcelDataValidation.List("C1:C10", ["bajo", "medio", "alto"]));
    sheet.AutoFilter = new ExcelAutoFilter("D1:E10");
    sheet.AddTable("Resumen", "A1:C10");

    // La vista y la impresión de la hoja.
    sheet.FreezePanes = new ExcelFreezePanes(rows: 1, columns: 0);
    sheet.SetColumnWidth(1, 24);
    sheet.PageSetup = new ExcelPageSetup(ExcelPageOrientation.Landscape, ExcelPaperSize.A4);

    // Un estilo reutilizable agrupa fuente, relleno, bordes y alineación. El formato numérico es
    // independiente y se conserva al aplicar el estilo.
    CellStyle headerStyle = new()
    {
        Bold = true,
        FontColor = RgbColor.FromRgb(0xFF, 0xFF, 0xFF),
        FillColor = RgbColor.FromRgb(0x1F, 0x49, 0x7D),
        Borders = new CellBorders(bottom: new CellBorder(CellBorderStyle.Thin)),
        HorizontalAlignment = CellHorizontalAlignment.Center,
    };
    sheet["A1"].Style = headerStyle;
    sheet["B1"].Style = headerStyle;

    workbook.Save("resumen.xlsx");
}

// PDF: convertir a PDF con LibreOffice en modo headless. El motor va detrás de `IPdfConverter`,
// así que el núcleo no depende de él. La conversión usa los bytes que se le dan: aquí, el estado
// actual del documento, guardado en memoria.
using (WordDocument document = WordDocument.Create())
{
    document.AddParagraph("Informe mensual");

    using MemoryStream current = new();
    document.Save(current);
    current.Position = 0;

    using LibreOfficePdfConverter converter = new();
    PdfConversionResult result = converter.Convert(current, PdfSourceFormat.WordDocument);
    File.WriteAllBytes("informe.pdf", result.Pdf);
}
```

Errores tipados:

```csharp
try
{
    using WordDocument document = WordDocument.Open("entrada.docx");
}
catch (InvalidOfficeFileException ex)
{
    // El archivo no es un DOCX legible (incluye el caso de archivos cifrados).
    Console.WriteLine(ex.Message);
}
catch (OfficeResourceLimitExceededException ex)
{
    // El archivo supera los límites configurados.
    Console.WriteLine(ex.Message);
}
```

> Nota: asignar una fórmula **no** está soportado todavía. El cálculo de fórmulas es un
> subsistema separado y no forma parte de las fases 1 y 2.

## Compilar y probar

```bash
dotnet build OfficeEngine.sln -c Release
dotnet test OfficeEngine.sln -c Release
```

La solución compila con **cero avisos** (`TreatWarningsAsErrors`) y la suite pasa al 100 %
(estado y cifras reales en [`docs/phase-1-report.md`](docs/phase-1-report.md),
[`docs/phase-2-report.md`](docs/phase-2-report.md) y [`docs/phase-3-report.md`](docs/phase-3-report.md)).

## Estructura del repositorio

```
src/OfficeEngine.Core      Contratos, opciones, límites, diagnósticos y errores tipados
src/OfficeEngine.OpenXml   Infraestructura OOXML interna (carga acotada + inspección de paquete)
src/OfficeEngine.Word      Modelo y edición DOCX
src/OfficeEngine.Excel     Modelo y edición XLSX
src/OfficeEngine.Pdf       Contratos de conversión a PDF
src/OfficeEngine.Pdf.LibreOffice  Adaptador de conversión con LibreOffice headless
src/OfficeEngine.Legacy    Conversión de los formatos heredados (.doc, .xls)
src/OfficeEngine.Formulas  Motor de cálculo de fórmulas
src/OfficeEngine.LibreOffice  Infraestructura interna del proceso externo (compartida)
src/OfficeEngine.Benchmarks Medición de rendimiento sin dependencias
tests/OfficeEngine.Tests   Pruebas (xUnit): unitarias, round-trip, interoperabilidad y seguridad
docs/                      Arquitectura, ADR, matriz de compatibilidad, riesgos y hoja de ruta
docs-site/                 Sitio de documentación (Docusaurus): índice, Word, Excel y PDF
```

## Documentación

- **Sitio de documentación (Docusaurus)**: [`docs-site/`](docs-site) — índice, Word, Excel, PDF y
  formatos heredados, con un ejemplo por función. Para verlo en local:

  ```bash
  cd docs-site
  npm install
  npm start
  ```

- Arquitectura: [`docs/architecture/overview.md`](docs/architecture/overview.md)
- Decisiones (ADR): [`docs/adr/`](docs/adr)
- Matriz de compatibilidad: [`docs/compatibility-matrix.md`](docs/compatibility-matrix.md)
- Riesgos, dependencias y licencias: [`docs/risks-and-dependencies.md`](docs/risks-and-dependencies.md)
- Hoja de ruta: [`docs/roadmap.md`](docs/roadmap.md)
- Informe de la Fase 1: [`docs/phase-1-report.md`](docs/phase-1-report.md)
- Informe de la Fase 2 (incrementos 2.1 a 2.12): [`docs/phase-2-report.md`](docs/phase-2-report.md)
- Informe de la Fase 3 (incrementos 3.1 a 3.11): [`docs/phase-3-report.md`](docs/phase-3-report.md)
- ADR del formato de párrafo: [`docs/adr/0009-paragraph-formatting-model.md`](docs/adr/0009-paragraph-formatting-model.md)
- ADR de tabulaciones y saltos: [`docs/adr/0010-tab-stops-and-breaks.md`](docs/adr/0010-tab-stops-and-breaks.md)
- ADR de referencias y catálogo de estilos: [`docs/adr/0011-style-references-and-catalogue.md`](docs/adr/0011-style-references-and-catalogue.md)
- ADR de listas (referencias, catálogo y creación): [`docs/adr/0012-lists-references-catalogue-and-creation.md`](docs/adr/0012-lists-references-catalogue-and-creation.md)
- ADR de tablas (estructura y contenido): [`docs/adr/0013-tables-structure-and-content.md`](docs/adr/0013-tables-structure-and-content.md)
- ADR de tablas (propiedades de tabla y de celda): [`docs/adr/0014-table-and-cell-properties.md`](docs/adr/0014-table-and-cell-properties.md)
- ADR de tablas (combinación de celdas): [`docs/adr/0015-cell-merge.md`](docs/adr/0015-cell-merge.md)
- ADR de secciones y configuración de página: [`docs/adr/0016-sections-and-page-setup.md`](docs/adr/0016-sections-and-page-setup.md)
- ADR de encabezados y pies: [`docs/adr/0017-headers-and-footers.md`](docs/adr/0017-headers-and-footers.md)
- ADR de imágenes y relaciones: [`docs/adr/0018-images-and-relationships.md`](docs/adr/0018-images-and-relationships.md)
- ADR de hipervínculos externos: [`docs/adr/0019-external-hyperlinks.md`](docs/adr/0019-external-hyperlinks.md)
- ADR de propiedades básicas del documento: [`docs/adr/0020-document-properties.md`](docs/adr/0020-document-properties.md)
- ADR del acceso a celdas por índice de fila/columna: [`docs/adr/0021-excel-cell-indices.md`](docs/adr/0021-excel-cell-indices.md)
- ADR de rangos de celdas: [`docs/adr/0022-excel-cell-ranges.md`](docs/adr/0022-excel-cell-ranges.md)
- ADR de fechas y formatos numéricos: [`docs/adr/0023-excel-dates-and-number-formats.md`](docs/adr/0023-excel-dates-and-number-formats.md)
- ADR de estilos reutilizables de celda: [`docs/adr/0024-excel-reusable-cell-styles.md`](docs/adr/0024-excel-reusable-cell-styles.md)
- ADR de escritura de fórmulas: [`docs/adr/0025-excel-formula-writing.md`](docs/adr/0025-excel-formula-writing.md)
- ADR de rangos con nombre: [`docs/adr/0026-excel-named-ranges.md`](docs/adr/0026-excel-named-ranges.md)
- ADR de validación de datos: [`docs/adr/0027-excel-data-validation.md`](docs/adr/0027-excel-data-validation.md)
- ADR de autofiltros: [`docs/adr/0028-excel-auto-filters.md`](docs/adr/0028-excel-auto-filters.md)
- ADR de tablas estructuradas: [`docs/adr/0029-excel-structured-tables.md`](docs/adr/0029-excel-structured-tables.md)
- ADR de paneles congelados, filas y columnas ocultas, alto y ancho: [`docs/adr/0030-excel-sheet-layout.md`](docs/adr/0030-excel-sheet-layout.md)
- ADR de configuración de impresión y página: [`docs/adr/0031-excel-print-settings.md`](docs/adr/0031-excel-print-settings.md)
- ADR de hipervínculos y notas: [`docs/adr/0032-excel-hyperlinks-and-comments.md`](docs/adr/0032-excel-hyperlinks-and-comments.md)
- ADR de imágenes en la hoja: [`docs/adr/0033-excel-images.md`](docs/adr/0033-excel-images.md)
- ADR de gráficos: [`docs/adr/0034-excel-charts.md`](docs/adr/0034-excel-charts.md)
- ADR de tabla de cadenas compartidas: [`docs/adr/0035-excel-shared-strings.md`](docs/adr/0035-excel-shared-strings.md)
- Informe de la Fase 4 (incrementos 4.1 y 4.2): [`docs/phase-4-report.md`](docs/phase-4-report.md)
- Informe de la Fase 5 (incrementos 5.1 a 5.4): [docs/phase-5-report.md](docs/phase-5-report.md)
- Informe de la Fase 6 (incrementos 6.1 y 6.2): [docs/phase-6-report.md](docs/phase-6-report.md)
- ADR de búsqueda y reemplazo en Word: [`docs/adr/0036-word-text-search-and-replace.md`](docs/adr/0036-word-text-search-and-replace.md)
- ADR de campos y numeración de páginas en Word: [`docs/adr/0037-word-fields-and-page-numbering.md`](docs/adr/0037-word-fields-and-page-numbering.md)
- ADR de conversión a PDF con LibreOffice headless: [`docs/adr/0038-pdf-libreoffice-converter.md`](docs/adr/0038-pdf-libreoffice-converter.md)
- ADR de carga de OOXML hostil y corpus de archivos corruptos: [`docs/adr/0039-hostile-ooxml-loading.md`](docs/adr/0039-hostile-ooxml-loading.md)
- ADR de modelo de almacenamiento en memoria y API síncrona: [`docs/adr/0040-in-memory-storage-and-synchronous-api.md`](docs/adr/0040-in-memory-storage-and-synchronous-api.md)
- ADR de medición de rendimiento y regresión de round-trip: [`docs/adr/0041-benchmarks-and-round-trip-regression.md`](docs/adr/0041-benchmarks-and-round-trip-regression.md)
- ADR de conversión de los formatos heredados: [docs/adr/0042-legacy-format-conversion.md](docs/adr/0042-legacy-format-conversion.md)
- ADR del motor de fórmulas: [docs/adr/0043-formula-engine.md](docs/adr/0043-formula-engine.md)
- Matriz de fórmulas soportadas: [docs/formula-matrix.md](docs/formula-matrix.md)
- Reglas de desarrollo: [`AGENT.md`](AGENT.md) · Especificación funcional: [`PROMPT.md`](PROMPT.md)

## Dependencias

- [`DocumentFormat.OpenXml`](https://www.nuget.org/packages/DocumentFormat.OpenXml) 3.5.1 — licencia
  MIT, mantenida por Microsoft. Se usa para el contenedor OPC y los modelos de esquema; **no**
  compone páginas ni genera PDF.

## Licencia

MIT. Ver [`LICENSE`](LICENSE).