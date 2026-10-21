# OfficeEngine — Fase 0: Arquitectura y decisiones

Este documento resume las decisiones tomadas en la Fase 0 y la estructura de la solución.
Las decisiones individuales están registradas como ADR en `docs/adr/`.

> Estado real: la Fase 1 (base compilable) está implementada y probada, la Fase 2 ha empezado
> con los incrementos 2.1 (runs y formato de carácter), 2.2 (alineación, sangría y espaciado de
> párrafo), 2.2b (tabulaciones y saltos), 2.3 (referencias de estilo y catálogo de solo lectura),
> 2.4 (listas: referencias, catálogo de solo lectura y creación), 2.5a (tablas: estructura y
> contenido), 2.5b (tablas: propiedades de tabla y de celda), 2.5c (tablas: combinación de celdas),
> 2.6 (secciones, tamaño/orientación y márgenes), 2.7 (encabezados y pies por sección), 2.8
> (imágenes inline y relaciones), 2.9 (hipervínculos externos), 2.10 (propiedades básicas del
> documento), 2.11 (búsqueda y reemplazo) y 2.12 (campos y numeración de páginas). La Fase 3 ha empezado con los incrementos 3.1 (acceso a celdas por índice de fila y
> columna en Excel), 3.2 (rangos de celdas), 3.3 (fechas, horas y formatos numéricos), 3.4 (estilos
> reutilizables de celda), 3.5 (escritura de fórmulas sin cálculo), 3.6a–3.6d (rangos con nombre,
> validación de datos, autofiltros y tablas estructuradas), 3.7 (paneles congelados, filas y columnas
> ocultas, alto y ancho), 3.8 (configuración de impresión y página), 3.9 (hipervínculos y notas),
> 3.10a–3.10b (imágenes y gráficos) y 3.11 (tabla de cadenas compartidas). La Fase 6 está entregada con los incrementos 6.1 (conversión de los formatos heredados con un motor externo) y 6.2 (motor de fórmulas propio, con la matriz publicada en `docs/formula-matrix.md`). La Fase 5 (robustez) está entregada con los incrementos 5.1 (carga de OOXML hostil), 5.2 (corpus de corruptos), 5.3 (documentos grandes y decisión de almacenamiento) y 5.4 (benchmarks y regresión de round-trip). La Fase 4 (conversión a
> PDF) está entregada con los incrementos 4.1 (contratos) y 4.2 (adaptador de LibreOffice headless).
> Todo lo demás descrito aquí como "planificado" no existe todavía y no debe considerarse
> soportado. La fuente de verdad del estado funcional es `docs/compatibility-matrix.md`.

## 1. Objetivo y alcance

OfficeEngine es una librería .NET independiente para **crear, leer, editar, guardar y
convertir** documentos Word y libros Excel. Prioridades (según `PROMPT.md` y `AGENT.md`):

1. Integridad y conservación de datos.
2. Seguridad frente a archivos no confiables.
3. Correctitud y compatibilidad comprobable con pruebas.
4. API clara y estable.
5. Rendimiento medido.
6. Amplitud de funcionalidades.

Formatos nativos iniciales: **DOCX** y **XLSX** (Office Open XML). DOC/XLS y la conversión
PDF se abordan como módulos separados en fases posteriores.

Spire.Doc y Spire.XLS se usan **solo** como referencia de categorías de funcionalidad y
experiencia de uso. No se copia código, recursos ni documentación propietaria, y no se
afirma paridad total.

## 2. Principios de arquitectura

| Decisión | Motivo |
| --- | --- |
| Núcleo sin motor de renderizado | La composición de páginas no es OOXML; mezclarla con el núcleo impide verificar y sustituir motores. |
| Cálculo de fórmulas como subsistema aparte | Escribir una fórmula no es calcularla. Mezclarlos genera promesas imposibles de probar. |
| DOC/XLS como módulo posterior | Son formatos binarios distintos; prometer lectura nativa al inicio sería falso. |
| Sin exponer tipos del SDK OOXML | La API pública es un contrato estable; los detalles internos deben poder cambiar. |
| Sin pérdida silenciosa | El contenido no modelado se preserva o se reporta por diagnósticos; nunca se descarta en silencio. |
| I/O síncrona en la primera etapa | Evitar una API `async` sobre un almacén en memoria que no lo es (ADR 0004). |
| Límites de recursos en la frontera pública | Los archivos Office son entradas no confiables (ADR 0007). |

## 3. Estructura de la solución

```
OfficeEngine.sln
├── src/
│   ├── OfficeEngine.Core            Contratos, opciones, límites, diagnósticos y errores tipados.
│   ├── OfficeEngine.OpenXml         Infraestructura interna OOXML (carga acotada + inspección de paquete).
│   ├── OfficeEngine.Word            Modelo y edición DOCX (WordprocessingML).
│   ├── OfficeEngine.Excel           Modelo y edición XLSX (SpreadsheetML).
│   ├── OfficeEngine.Pdf             Contratos de conversión a PDF (sin motor).
│   ├── OfficeEngine.Pdf.LibreOffice Adaptador de conversión con LibreOffice headless.
│   ├── OfficeEngine.Legacy          Conversión de los formatos heredados con un motor externo.
│   ├── OfficeEngine.LibreOffice     Infraestructura interna: ejecuta el motor de forma aislada.
│   ├── OfficeEngine.Formulas         Motor de cálculo de fórmulas, sin depender de Word ni de Excel.
│   └── OfficeEngine.Benchmarks      Medición del rendimiento, sin dependencias propias.
├── tests/
│   └── OfficeEngine.Tests           xUnit: unitarias, round-trip, interoperabilidad, PDF y seguridad.
├── docs/                            Arquitectura, ADR, matriz de compatibilidad, riesgos y hoja de ruta.
├── docs-site/                       Sitio de documentación (Docusaurus): índice, Word, Excel y PDF.
└── .github/workflows/ci.yml         Integración continua.
```

Proyectos **planificados pero no creados** (se añadirán cuando su fase empiece, no antes,
para no llenar la solución de contenedores vacíos):

| Proyecto | Fase | Propósito |
| --- | --- | --- |
| `OfficeEngine.Samples` | 7 | Ejemplos ejecutables contra la API real. |

Las dependencias entre proyectos son unidireccionales:

```
Core  ←  OpenXml  ←  Word
                   ←  Excel

Core  ←  Pdf  ←  Pdf.LibreOffice
```

`OfficeEngine.Pdf` no lo referencia ni `Word` ni `Excel`: la conversión se pide explícitamente y
con los bytes que da el llamante, de modo que el núcleo no depende de ningún motor.

`OfficeEngine.OpenXml` no es una superficie pública estable: sus tipos son `internal` y solo
son visibles para `Word`, `Excel` y los tests mediante `InternalsVisibleTo`.

## 4. Responsabilidades por capa

- **Core**: `OfficeEngineException` y derivadas tipadas, `OfficeEngineLimits`,
  `OfficeEngineDiagnostic`/`OfficeEngineDiagnostics`, `RgbColor`. No conoce OOXML.
- **OpenXml**: carga acotada de entradas no confiables (`OoxmlLoader`), inspección del contenedor ZIP
  antes de parsear (`OoxmlPackageInspector`) y **medida de las partes XML antes de que las parsee el
  SDK** (`OoxmlXmlGuard`), que rechaza los DTD y el anidamiento por encima del límite y evita que
  un archivo hostil termine el proceso. `OoxmlPartReader` traduce las excepciones del SDK y del
  lector de XML a los errores tipados de la librería. No conoce Word ni Excel.
- **Word**: modelo de dominio y edición sobre el paquete OOXML, cerrado detrás
  de una API pública con nombres familiares (`WordDocument`, `Paragraph`, `Run`, `UnderlineStyle`).
  Word modela texto y formato de carácter por run
  (incluyendo el orden de `w:rPr` y la semántica tri-state), formato de párrafo por párrafo
  (alineación, sangría, espaciado y tabulaciones, insertando los hijos de `w:pPr` en orden de
  esquema), el contenido de un run (texto, tabulación y saltos, expuesto por `Run.Content` como
  `RunText`/`RunTab`/`RunBreak`) y las **referencias de estilo** (`Paragraph.StyleId`,
  `Run.StyleId`), además de un **catálogo de estilos de solo lectura** (`WordDocument.Styles`,
  `FindStyle`, `WordStyle` con `OwnFormatting`/`EffectiveFormatting` y la herencia `w:basedOn`
  resuelta). No crea ni edita definiciones de estilo: la parte `styles.xml` se preserva. Word modela
  también las **listas**: la referencia de un párrafo a una lista y a un nivel (`Paragraph.ListId`,
  `Paragraph.ListLevel`, `Paragraph.ApplyList`, sobre `w:numPr`), un **catálogo de listas de solo
  lectura** (`WordDocument.Lists`, `FindList`, `WordList`/`WordListLevel` con formato, texto, inicio
  y justificación) y la **creación de listas de un solo nivel** (`WordDocument.AddList`), que crea
  `word/numbering.xml` cuando hace falta. No edita definiciones de lista existentes: el
  `numbering.xml` cargado se preserva. Word modela además las **tablas** en su estructura, su
  contenido, sus **propiedades** y su **combinación de celdas**: `WordDocument.Tables`/`AddTable`, la
  jerarquía `Table`/`TableRow`/`TableCell`, el texto de cada celda (`TableCell.Text`), el formato de
  la tabla y de la celda (ancho, alineación, disposición, bordes, sombreado y alineación vertical),
  el span y la combinación vertical de una celda (`TableCell.ColumnSpan`/`VerticalMerge`) y las
  operaciones de combinar y dividir (`TableRow.MergeCells`/`UnmergeCell`,
  `Table.MergeCellsVertically`/`UnmergeCellsVertically`), con la tabla como elemento de bloque al
  mismo nivel que el párrafo. Word modela las **secciones** (`WordDocument.Sections`/`AddSection`,
  `Section` con tamaño/orientación, márgenes y tipo de inicio) y sus **encabezados y pies**
  (`Section.Header`/`Footer`/`AddHeader`/`AddFooter`/`RemoveHeader`/`RemoveFooter`,
  `HeaderFooter` por tipo `default`/`first`/`even`, con párrafos y tablas editables y partes que
  varias secciones pueden compartir). Word modela también las **imágenes inline incrustadas**:
  `Paragraph.AddImage` desde archivo o stream (`ImageFormat`), `Paragraph.Images` e `Image` con
  tamaño en puntos, tipo de contenido y bytes, en el cuerpo, celdas de tabla y encabezados/pies; cada
  imagen crea su parte y relación. Word modela los **hipervínculos externos**:
  `Paragraph.AddHyperlink`, `Paragraph.Hyperlinks` y `Hyperlink` con texto, runs directos y destino
  editables, en el cuerpo, celdas de tabla y encabezados/pies; el destino vive en una relación de
  hipervínculo de la parte y editarlo no cambia otros enlaces que compartían esa relación. Word modela
  también las **propiedades básicas del documento** (`WordDocument.Properties`,
  `DocumentProperties`): título, asunto, autor, palabras clave, descripción, último editor,
  revisión, fechas, categoría, estado, identificador, idioma y versión, sobre `docProps/core.xml`
  (la parte se crea solo al asignar el primer valor). Word modela además la **búsqueda y el
  reemplazo** (`WordDocument.FindText`/`ReplaceText` y `Paragraph.FindText`/`ReplaceText`, con
  `WordTextMatch`): el alcance es el cuerpo, la búsqueda no cruza párrafos y el texto reemplazado
  conserva el formato del primer run de la coincidencia. Y los **campos** (`Paragraph.AddField`,
  `HeaderFooter.AddPageNumber`, `WordField`), que se escriben con la secuencia de runs de un campo y
  su resultado almacenado, sin evaluarlos. Lo que no
  se modela (márgenes de celda, `w:hMerge` heredado, las combinaciones rectangulares, columnas de
  sección, imágenes flotantes o vinculadas, gráficos, hipervínculos internos a marcadores, atributos
  de hipervínculo distintos del destino, contenido no modelado de encabezados/pies y propiedades
  extendidas o personalizadas del documento) y las tablas
  anidadas se preservan; el contenido que no modela se conserva sin cambios.

- **Excel**: modelo de dominio y edición del paquete SpreadsheetML, cerrado detrás de una API
  pública (`ExcelWorkbook`, `Worksheet`, `Cell`). Una celda se direcciona por su referencia A1 o por
  índices 1-based de fila y columna (`Worksheet.GetCell`/`Worksheet[int, int]`), y cada `Cell` expone
  `Reference`, `Row` y `Column`. Un rango (`Worksheet.GetRange`) expone sus dos esquinas, recorre sus
  celdas en orden de fila y lee o escribe una matriz con `GetValues`/`SetValues`; las celdas ausentes
  se leen como `null` y no se crean. Al escribir se admiten texto, número, booleano, fecha/hora y
  vacío. Una fecha se guarda como serial y necesita un formato explícito (`Cell.SetDateTime`);
  `Cell.NumberFormat` lee y escribe el código efectivo, y el formato se resuelve en el componente
  interno `ExcelStyles`, que crea un `xl/styles.xml` mínimo de forma perezosa y preserva el existente.
  `CellRange.GetValues` devuelve `DateTime` para celdas con formato de fecha/hora; `SetValues` no
  escribe `DateTime` porque cada celda necesita su propio formato. `Cell.Style` modela un estilo
  reutilizable (`CellStyle`: fuente, relleno sólido, cuatro bordes y alineación); aplicar el mismo
  estilo a varias celdas reutiliza `fonts`/`fills`/`borders`/`cellXfs`, y el formato numérico y las
  propiedades no modeladas de fuente, borde y alineación se conservan. `Cell.Formula` escribe
  fórmulas como fórmulas, sin calcular nada: al escribir se declara ausente el resultado almacenado,
  al leer se devuelve el que había, y el libro marca `calcPr/@fullCalcOnLoad` para que el motor que
  lo abra recalcule. `ExcelWorkbook.DefinedNames` lista los nombres definidos del libro (ámbitos de
  libro y de hoja) y `Worksheet.AddNamedRange` los crea; los nombres que no son rangos se conservan
  sin interpretar. `Worksheet` expone además las reglas de validación de datos
  (`AddDataValidation`), el autofiltro (`AutoFilter`) y las tablas estructuradas (`AddTable`,
  `Tables`), que crean su propia parte `xl/tables/tableN.xml` con su relación. `Worksheet` expone
  además la vista de la hoja (`FreezePanes`, alto y ancho de filas y columnas, y su visibilidad) y
  la configuración de impresión (`PageMargins`, `PageSetup`, `PrintOptions`), los hipervínculos
  (`AddHyperlink`, `AddInternalHyperlink`), las notas (`AddComment`), las imágenes (`AddImage`) y los
  gráficos (`AddChart`), que crean su propia parte de comentarios, el dibujo heredado, la parte de
  medios y la parte de gráfico. El texto de las celdas se guarda en la tabla de cadenas
  compartidas (`xl/sharedStrings.xml`), con deduplicación. El resto del
  contenido de la hoja
  (por ejemplo `mergeCells`) se preserva y se reporta por diagnósticos.

- **PDF**: conversión a PDF separada del núcleo. `OfficeEngine.Pdf` define el contrato
  (`IPdfConverter`, `PdfSourceFormat`, `PdfConversionOptions`, `PdfConversionResult` y
  `PdfConversionException`) sin depender de ningún motor; `OfficeEngine.Pdf.LibreOffice` lo implementa
  con LibreOffice en modo headless: localiza el ejecutable, comprueba su versión, convierte los bytes
  que recibe en un directorio temporal propio y con un perfil de usuario aislado, y valida el código
  de salida y el PDF producido. Ni `Word` ni `Excel` dependen de este contrato: el llamante guarda
  el documento en memoria y entrega ese flujo al conversor.

## 5. Modelo de ciclo de vida

- Un documento abierto mantiene el paquete OOXML completo **en memoria** (buffer propio).
- `Open(path)` copia el archivo a memoria antes de parsear: no se retiene un bloqueo de
  archivo, por lo que `Save` sobre el mismo archivo es seguro.
- `Open(stream, leaveOpen)` copia el stream y, salvo `leaveOpen: true`, lo cierra.
- `Save(path)` escribe en un archivo temporal del mismo directorio y luego lo mueve de
  forma atómica, para no dejar el destino a medio escribir si algo falla.
- `Dispose` libera el paquete y el buffer. Los tipos son *thread-affine*.

Consecuencia y limitación: el consumo de memoria es proporcional al tamaño del archivo. Un
almacén respaldado por archivo o por segmentos está planificado (ADR 0004).

## 6. Interoperabilidad y verificación

La implementación no se da por buena porque el ZIP se abra. Cada capacidad se verifica con otra
herramienta:

- `OpenXmlValidator` del SDK oficial valida los paquetes producidos.
- `System.IO.Compression` lee las entradas crudas (por ejemplo `word/document.xml`) para
  comprobar el contenido semántico.
- Fixtures creados directamente con el SDK se leen a través de OfficeEngine.

## 7. Riesgos principales

Ver `docs/risks-and-dependencies.md`. Los más relevantes a corto plazo:

| Riesgo | Mitigación actual |
| --- | --- |
| Pérdida de contenido no modelado | Se muta el paquete original en lugar de reconstruirlo; el contenido desconocido se reporta por diagnósticos y se verifica con round-trip. |
| Bomba ZIP / entradas maliciosas | Inspección del directorio central sin descomprimir + límites configurables. |
| Consumo de memoria proporcional al archivo | Límite de tamaño; almacén en streaming planificado (ADR 0004). |
| Deriva de dependencias | Central Package Management y versión fijada (ADR 0003). |
| Fidelidad de PDF imposible de garantizar | Motor externo tras una interfaz; se documenta que la salida depende del entorno (ADR 0005). |
| Compatibilidad con DOC/XLS | Fase 6, mediante conversión externa; nunca nativa al inicio (ADR 0006). |
