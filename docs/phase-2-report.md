# Fase 2 - Informe de resultados reales (incrementos 2.1 a 2.12)

Fecha: 2026-10-16 (incrementos 2.1, 2.2 y 2.2b), 2026-10-17 (incrementos 2.3, 2.4 y 2.5a),
2026-10-18 (incrementos 2.5b y 2.5c) y 2026-10-19 (incrementos 2.6, 2.7, 2.8, 2.9 y 2.10).
Entorno: Windows 11 Enterprise LTSC (10.0.26100, x64), .NET SDK 8.0.426.
Comandos ejecutados y resultados **observados** (no estimados).

> Este informe se amplía incremento a incremento. Las secciones siguientes describen primero el
> incremento 2.1 (formato de carácter), después el 2.2 (formato de párrafo), luego el 2.2b
> (tabulaciones y saltos), después el 2.3 (referencias de estilo y catálogo de solo lectura), luego
> el 2.4 (listas), el 2.5a (tablas: estructura y contenido), el 2.5b (tablas: propiedades de tabla y
> de celda), el 2.5c (tablas: combinación de celdas), el 2.6 (secciones, tamaño/orientación y
> márgenes), el 2.7 (encabezados y pies por sección), el 2.8 (imágenes y relaciones) y por último el
> 2.9 (hipervínculos externos) y el 2.10 (propiedades básicas del documento).

## Alcance del incremento 2.1

Párrafos/runs editables y **formato de carácter**: familia y tamaño de fuente, negrita, cursiva,
subrayado y color RGB. Incluye:

- `Paragraph.Runs` (vista viva, en orden), `Paragraph.AddRun(string)`.
- `Run.Text` editable sin perder el formato del propio run.
- `Run.FontFamily`, `Run.FontSizePoints`, `Run.Bold`, `Run.Italic`, `Run.Underline`, `Run.Color`
  con semántica tri-state donde el OOXML la tiene.
- Nuevo primitivo compartido `OfficeEngine.RgbColor` (estructura de solo lectura).

## Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Debug` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release` | **122 superadas, 0 con error, 0 omitidas** (122 total). Duración: ~200 ms. |

Salida textual del resumen de `dotnet test`:

```
Correctas! - Con error: 0, Superado: 122, Omitido: 0, Total: 122,
Duración: 203 ms - OfficeEngine.Tests.dll (net8.0)
```

El informe TRX se generó en `artifacts/testresults/phase2/phase2.trx` y la cobertura en
`artifacts/testresults/phase2/**/coverage.cobertura.xml`.

## Cobertura de líneas (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % |
| --- | --- | --- |
| `Core/RgbColor.cs` | 31 / 31 | 100 % |
| `Word/Paragraph.cs` | 31 / 31 | 100 % |
| `Word/Run.cs` | 200 / 200 | 100 % |
| **Total de la solución (todos los archivos)** | 787 / 867 | 90,8 % |

Se cubrieron al 100 % los archivos nuevos de este incremento. El total de la solución incluye
código de Fase 1 no tocado por estas pruebas. Las cifras proceden del archivo Cobertura de esa
ejecución: el recuento absoluto de líneas por archivo no es comparable entre ejecuciones (véase la
nota del incremento 2.2).

## Pruebas nuevas del incremento (36 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordCharacterFormatTests.cs` (19) | API de runs y formato: orden, reemplazo de texto sin perder formato, round-trip por archivo y stream, precisión de medio punto, tri-state (heredar/desactivado/activado), los 18 subrayados, limpieza de propiedades, argumentos inválidos, `ToString`, vistas vivas. |
| `Word/WordFormattingInteropTests.cs` (10) | Validación de esquema del paquete producido, lectura con el SDK, orden de esquema de `w:rPr`, lectura de un archivo ajeno, preservación de `w:rPr` no modelado, color de tema y subrayado desconocido, y tres casos de archivo ilegible. |
| `Core/RgbColorTests.cs` (7) | Constructor, hex, `TryParseHex`, igualdad y `ToString` del color. |

## Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` sobre el paquete producido con formato
   (`ProducedDocument_WithCharacterFormatting_IsSchemaValid`).
2. **Lectura con el SDK** de los `w:rPr` producidos, incluyendo el caso de negrita desactivada
   explícitamente, que Word distingue de «sin propiedad».
3. **Lectura cruda** de `word/document.xml` con `System.IO.Compression` para comprobar el orden
   de los hijos de `w:rPr` (`ProducedDocument_WritesRunPropertiesInSchemaOrder`).
4. **Fixtures creados con el SDK** y leídos a través de OfficeEngine: formato explícito, color de
   tema, y contenido ilegible inyectado como XML crudo.

## Hallazgos durante las pruebas y acciones tomadas

1. **El SDK filtra una excepción de la BCL.** Al leer un `ST_Underline` o un `ST_OnOff`
   desconocido, el SDK lanza `System.FormatException`. OfficeEngine la convierte en
   `InvalidOfficeFileException` (`Run.ReadStoredValue`), cumpliendo el contrato de errores
   tipados. La prueba del subrayado desconocido lo demuestra.
2. **Comportamiento del SDK verificado empíricamente.** Se comprobó, con el SDK 3.5.1, que la
   lectura de valores on/off o de enumeración malformados lanza **siempre** `FormatException` (no
   `ArgumentOutOfRangeException`, que solo aparece en la ruta de construcción/validación). Por eso
   se captura únicamente `FormatException`: no se deja código defensivo no probado.
3. **Mutación parcial evitada.** La asignación de un subrayado rechazado creaba el elemento
   `<w:u>` antes de validar. Se reordenó para validar **antes** de tocar el árbol; hay prueba de
   que un valor inválido no deja marcado alguno
   (`Run_Underline_RejectsAValueOutsideTheModelledSetWithoutTouchingTheDocument`).
4. **Analizadores.** Se atendieron CA1720 (miembros `Single`/`Double` → `SingleLine`/`DoubleLine`),
   CA1859 (usar `Dictionary` concreto en el mapa de subrayados) y CA1861 (matriz constante).
5. **Fixtures inválidos.** Un `w:color` colocado directamente bajo `w:r` y sin `@val` producía un
   paquete inválido; se corrigieron **los fixtures**, no la librería. La validación con el SDK
   volvió a detectar un error real de esquema.

## Qué **no** se ha implementado (y por tanto no se afirma)

Ver `docs/compatibility-matrix.md`. En este incremento, en particular:

- No se modelan **resaltado, tachado, sombreados ni colores de tema**; se **preservan** tal cual
  y hay prueba de ello.
- `Run.Color` devuelve `null` para un color de tema o `auto`, porque no tienen RGB fijo.
- No se modelan fuentes ni tamaños de escritura compleja (`w:rFonts` `eastAsia`/`cs`, `w:szCs`).
- En 2.1 no había formato de párrafo (se añadió en 2.2); siguen sin modelarse estilos, listas,
  tablas editables, secciones, encabezados/pies, imágenes, hipervínculos ni búsqueda/reemplazo.

## Limitaciones conocidas

- Documento completo en memoria; máximo por `OfficeEngineLimits.MaxFileSizeBytes` (256 MiB).
- I/O síncrona (ADR 0004).
- Los diagnósticos se calculan al abrir y no se recalculan tras editar.
- El modelo de tres estados, la precisión de medio punto y la frontera de errores se documentan
  en `docs/adr/0008-character-formatting-model.md`.

## Incremento 2.2 — Formato de párrafo

### Alcance

Alineación, sangrías y espaciado de párrafo. Incluye:

- `Paragraph.Alignment` (`ParagraphAlignment?`, con los cinco valores modelados de `ST_Jc`).
- `Paragraph.FirstLineIndentPoints`, `LeftIndentPoints`, `RightIndentPoints` y
  `HangingIndentPoints` (`double?`, en puntos, sobre `w:ind`).
- `Paragraph.SpacingBeforePoints`, `SpacingAfterPoints` y `LineSpacingPoints` (`double?`, en
  puntos, sobre `w:spacing`).
- Inserción de las propiedades de párrafo en el orden de esquema de `CT_PPr`, conservando el
  marcado no modelado que ya estuviera presente.
- Lectura explícita (error tipado) de valores que la versión no modela, en lugar de valores
  inventados.

Las **tabulaciones** (`w:tabs`) y los **saltos** se separaron al incremento **2.2b** para no ampliar
el alcance a mitad de camino; se refleja en `docs/roadmap.md` y en la matriz.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Debug` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Debug` | **150 superadas, 0 con error, 0 omitidas** (150 total). Duración: 164 ms. |
| `dotnet test OfficeEngine.sln -c Release` | **150 superadas, 0 con error, 0 omitidas** (150 total). Duración: 167 ms. |

Salida textual del resumen de `dotnet test` (Release):

```
Correctas! - Con error: 0, Superado: 150, Omitido: 0, Total: 150,
Duración: 167 ms - OfficeEngine.Tests.dll (net8.0)
```

Informes TRX: `artifacts/testresults/phase2/inc22/phase2_inc22.trx` y
`artifacts/testresults/phase2/inc22/phase2_inc22_release.trx`. Cobertura en
`artifacts/testresults/phase2/inc22/coverage/**/coverage.cobertura.xml`.

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Word/Paragraph.cs` | 314 / 314 | 100 % | 91,1 % |
| `Word/Run.cs` | 277 / 277 | 100 % | 90,2 % |
| `Core/RgbColor.cs` | 38 / 38 | 100 % | 90,0 % |
| **Total de la solución** | 1356 / 1481 | 91,6 % | 82,0 % |

Las cifras por archivo se toman del bloque `<lines>` de cada clase en el archivo Cobertura de
**esta** ejecución; el recuento absoluto solo es comparable dentro de una misma ejecución.
`Paragraph.cs` pasó de ser un tipo pequeño a 314 líneas contadas. Las ramas no cubiertas son rutas
de validación redundantes (combinaciones de `NaN`/infinito por propiedad) y la rama de `RankOf`
para un hijo de `w:pPr` que no está en la tabla (que solo puede venir de XML no identificado), no
caminos funcionales.

### Pruebas nuevas del incremento (28 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordParagraphFormatTests.cs` (22) | Valores por defecto y `null`, round-trip por stream, los cinco valores de alineación, sangrías y espaciados, validación de signo y finitud, cambio de alineación sin duplicar `<w:jc>`, limpieza que elimina `w:pPr`/`w:ind`/`w:spacing` vacíos, conservación de atributos no modelados (`w:leftChars`), valores no modelados o mal formados con error tipado. |
| `Word/WordParagraphFormatInteropTests.cs` (6) | Validación de esquema, lectura con el SDK de `w:jc`/`w:ind`/`w:spacing`/`w:lineRule`, orden de esquema al aplicar en orden inverso, preservación de `w:keepNext` al editar, precedencia de `w:numPr` y de las banderas de párrafo no modeladas (`w:bidi`, `w:snapToGrid`…). |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` sobre los paquetes producidos con formato de
   párrafo (`ProducedDocument_WithParagraphFormatting_IsSchemaValid`).
2. **Lectura con el SDK** de `w:jc`, `w:ind` y `w:spacing`, comprobando los valores en twips y que
   se escribe `w:lineRule="exact"`.
3. **Lectura cruda** de `word/document.xml` para el orden de los hijos de `w:pPr`, aplicando las
   propiedades en orden inverso al del esquema.
4. **Fixtures creados con el SDK** (XML crudo y elementos del SDK) leídos por OfficeEngine: valor
   no modelado (`mediumKashida`), valor mal formado, `w:lineRule="auto"`, `w:leftChars` y `w:numPr`.

### Hallazgos durante las pruebas y acciones tomadas

1. **`w:firstLine` y `w:hanging` no admiten negativos.** El validador del SDK rechaza
   `w:hanging="-120"` y `w:firstLine="-120"` (`ST_TwipsMeasure` sin signo), mientras que `w:left`
   sí admite negativos. La API aceptaba negativos para la sangría francesa; se corrigió el contrato
   (`ArgumentOutOfRangeException`) y las pruebas, **en lugar de** escribir paquetes inválidos.
2. **`w:line` no siempre está en puntos.** Con `w:lineRule="auto"` es un múltiplo de línea; devolver
   `LineSpacingPoints` como puntos habría sido un número incorrecto. Ahora la lectura lanza
   `UnsupportedOfficeFeatureException` y la escritura fija `w:lineRule="exact"`.
3. **Duplicado de `w:jc`.** La asignación de la alineación añadía un segundo elemento en vez de
   reemplazarlo. Se corrigió con `InsertProperty` (que desengancha y reinserta) y hay prueba de que
   solo queda un `w:jc`.
4. **Tabla de orden con un nombre incorrecto.** La tabla usaba el nombre de la propiedad del SDK
   (`numberingProperties`) en lugar del nombre local XML (`numPr`), de modo que `w:numPr` quedaba
   desordenado. Se detectó al escribir la prueba de numeración y se corrigió.
5. **`OpenXmlElement.Remove()` lanza si el elemento no tiene padre.** El primer intento de
   `InsertProperty` falló en las pruebas; ahora solo se desengancha cuando hay padre.
6. **Desbordamiento silencioso al convertir puntos a twips.** Un valor finito enorme
   (`double.MaxValue`) pasaba la validación y `(int)Math.Round(puntos * 20)` producía
   `-2147483648` (se comprobó con un programa mínimo: `(int)` de un `double` fuera de rango no
   lanza, devuelve un valor sin sentido). Escribir eso habría sido exactamente el «resultado
   incorrecto en silencio» que el proyecto prohíbe. `PointsToTwips` ahora comprueba el rango de
   `Int32` y lanza `ArgumentOutOfRangeException`; la prueba se añadió **antes** del arreglo y falló
   primero (rojo → verde).
7. **La tabla de orden de `w:pPr` estaba incompleta y con nombres equivocados.** Contenía
   `overflowPunctuation` y `topLinePunctuation` (los nombres reales son `overflowPunct` y
   `topLinePunct`), un `kin` inexistente y omitía `bidi`, `adjustRightInd` y `snapToGrid`. Un
   párrafo ajeno con `w:bidi`/`w:snapToGrid` habría recibido `w:spacing` **antes** que esas banderas,
   es decir, un paquete inválido. La secuencia se volvió a derivar del orden en que el SDK genera los
   hijos de `ParagraphProperties` (sonda de reflexión sobre 3.5.1) y se añadió la prueba
   `OpenedDocument_WithUnmodelledParagraphFlags_KeepsSchemaOrder`, que falló primero (rojo → verde).

### Qué **no** se ha implementado (y por tanto no se afirma)

- Tabulaciones (`w:tabs`) y saltos de línea/página/columna: se entregaron en el incremento 2.2b
  (sección siguiente).
- La sangría francesa negativa y la sangría en caracteres (`w:leftChars`, `w:firstLineChars`) no se
  modelan; la última se **preserva**.
- La decisión de diseño (unidades, orden de esquema, valores no modelados) está en
  `docs/adr/0009-paragraph-formatting-model.md`.

## Incremento 2.2b — Tabulaciones y saltos

### Alcance

Tabulaciones de párrafo y el contenido no textual de un run (tabuladores y saltos). Incluye:

- `TabStop` (tipo de valor: `PositionPoints`, `TabStopAlignment`, `TabStopLeader`) y
  `Paragraph.TabStops`, `Paragraph.AddTabStop(TabStop)` y `Paragraph.ClearTabStops()`, sobre
  `w:pPr/w:tabs`. La posición se interpreta en puntos (con signo, como el esquema) y se escribe
  redondeada a veinteavos de punto.
- `TabStopAlignment` (`Clear`, `Left`, `Center`, `Right`, `Decimal`, `Bar`) y `TabStopLeader`
  (`None`, `Dot`, `Hyphen`, `Underscore`, `Heavy`, `MiddleDot`).
- `BreakType` (`TextWrapping`, `Page`, `Column`) y la vista fiel `Run.Content`, compuesta por
  `RunContent` (`RunText`, `RunTab`, `RunBreak`), sobre `w:tab`, `w:br` y `w:cr`.
- `Run.AddTab()` y `Run.AddBreak(BreakType)`.
- Rechazo tipado de tabuladores y saltos (y de caracteres no representables en XML) en todas las
  entradas de texto plano, para no escribirlos «como texto».

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Debug` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Debug` | **207 superadas, 0 con error, 0 omitidas** (207 total). |
| `dotnet test OfficeEngine.sln -c Release` | **207 superadas, 0 con error, 0 omitidas** (207 total). |

Salida textual del resumen de `dotnet test` (Release, la duración varía algunas decenas de
milisegundos entre ejecuciones):

```
Correctas! - Con error: 0, Superado: 207, Omitido: 0, Total: 207,
Duración: 228 ms - OfficeEngine.Tests.dll (net8.0)
```

Informes TRX: `artifacts/testresults/phase2/inc22b/coverage/phase2_inc22b.trx` y
`artifacts/testresults/phase2/inc22b/phase2_inc22b_release.trx`. Cobertura en
`artifacts/testresults/phase2/inc22b/coverage/**/coverage.cobertura.xml`.

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Word/Paragraph.cs` | 453 / 453 | 100 % | 92,8 % |
| `Word/Run.cs` | 340 / 342 | 99,4 % | 91,2 % |
| `Word/PlainText.cs` | 24 / 24 | 100 % | 100 % |
| `Word/TabStop.cs` | 20 / 20 | 100 % | 87,5 % |
| `Word/RunContent.cs`, `RunText.cs`, `RunTab.cs`, `RunBreak.cs` | 13 / 13 | 100 % | — |
| **Total de la solución** | 1616 / 1743 | 92,7 % | 84,1 % |

Las dos únicas líneas sin cubrir de `Run.cs` son el `throw` de la rama de `ReadBreakType` para un
tipo de salto que el SDK conoce pero OfficeEngine no modela: `ST_BrType` solo define `textWrapping`,
`page` y `column`, que sí se modelan, de modo que la rama es defensiva y hoy inalcanzable (se
mantiene por coherencia con el resto de mapeos). Las ramas parciales que quedan son fallbacks
equivalentes (mapeos de enumeración ya exhaustivos y comprobaciones de `Equals`), no caminos
funcionales.

### Pruebas nuevas del incremento (57 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordTabStopTests.cs` (23) | Estado por defecto, añadir/leer, redondeo a twip, posición negativa, round-trip por stream, `ClearTabStops` (incluida la poda de `w:pPr` vacío y su ausencia), las 36 combinaciones de alineación y relleno escritas y releídas, conservación del resto del formato, omisión de `w:leader` con `None`, posición fuera de rango sin dejar marcado, alineación/relleno no definidos, lectura de las seis alineaciones y los seis rellenos, y casos de archivo ilegible (sin `w:val`, sin `w:pos`, posición o alineación mal formadas, relleno mal formado) y de valor válido no modelado (`num`), preservado. |
| `Word/WordRunContentTests.cs` (27) | `Run.Content` en orden, `Run.Text` como texto plano, `AddTab`/`AddBreak` (los tres tipos), rechazo de un tipo no definido, reemplazo del contenido por `Run.Text`, round-trip, conservación de espacios, `w:cr` ajeno leído como salto de línea, cada `w:type` almacenado (incluida la omisión y `textWrapping`), tipo de salto mal formado, conservación de `w:cr`/`w:lastRenderedPageBreak`, rechazo de tabuladores y saltos en las seis entradas de texto, rechazo de caracteres no representables (`U+0000`, `U+0007`) y de suplentes sueltos, y aceptación de un par de suplentes válido. |
| `Word/WordTabAndBreakInteropTests.cs` (7) | Validación de esquema del paquete producido, lectura con el SDK de `w:tabs` (posición en twips, alineación, relleno) y de `w:tab`/`w:br`, orden de esquema de `w:pPr` (`w:keepNext`, `w:tabs`, `w:ind`, `w:jc`), ausencia de `w:type="textWrapping"`, conservación del marcado no modelado al añadir una parada, y round-trip de un archivo ajeno con paradas previas. |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes producidos con paradas y saltos
   (`ProducedDocument_WithTabsAndBreaks_IsSchemaValid`), al añadir una parada a un párrafo ajeno
   (`AddingATabStop_KeepsUnmodelledParagraphMarkupInSchemaOrder`), al escribir las 36 combinaciones
   de alineación y relleno, y del paquete con `w:cr`/`w:lastRenderedPageBreak` conservado.
2. **Lectura con el SDK** de `w:tabs`/`w:tab`/`w:br`, comprobando posición (twips), alineación y
   relleno, el número de tabuladores y saltos, y el `w:type` de cada salto.
3. **Lectura cruda** de `word/document.xml` para el orden de los hijos de `w:pPr` (con `w:tabs`
   entre `w:keepNext` y `w:ind`/`w:jc`) y para comprobar que no se escribe un
   `w:type="textWrapping"` explícito.
4. **Fixtures creados con el SDK** (XML crudo y elementos del SDK) leídos por OfficeEngine: tabulaciones
   con cada alineación/relleno, paradas sin atributos obligatorios o mal formadas, alineación no
   modelada (`num`), saltos con cada `w:type`, salto sin `w:type` y salto mal formado; contenido mixto
   con `w:cr` y `w:lastRenderedPageBreak`.

### Hallazgos durante las pruebas y acciones tomadas

1. **Un tabulador o un salto como carácter no era un error.** `w:t` admite `\t` y `\n` y el paquete
   se guarda sin avisar (comprobado con la sonda), pero se renderiza distinto de `w:tab`/`w:br`. Se
   añadió `PlainText.EnsureValid` en todas las entradas de texto y se documentó que los tabuladores
   y saltos se construyen con `Run.AddTab`/`Run.AddBreak`.
2. **Caracteres que el SDK no puede serializar fallaban tarde y con una excepción cruda.** Un `\u0000`
   o un suplente suelto pasaban la validación de argumentos y hacían fallar `Save()` con una
   `ArgumentException` de la BCL (comprobado con un programa mínimo). Ahora se rechazan al asignar,
   con el parámetro correcto, y se aceptan los pares de suplentes bien formados.
3. **Un `w:tab` de párrafo mal formado no lanzaba `FormatException` en el punto esperado.** Asignar un
   valor de enumeración inválido a una propiedad tipada del SDK **valida inmediatamente** (la
   conversión implícita a `EnumValue<T>` construye y valida), de modo que un fixture «realista» hay
   que construirlo con `SetAttribute`/XML crudo. Los fixtures se reescribieron así y se comprobó que
   la lectura ejercita de verdad la ruta `FormatException → InvalidOfficeFileException`.
4. **La posición de una tabulación es con signo.** `w:pos` es `ST_SignedTwipsMeasure`; a diferencia de
   las sangrías (que rechazan el signo negativo, ADR 0009), aquí se acepta y se reproduce, porque el
   esquema lo permite de forma explícita. Se decidió y documentó en el ADR 0010; hay prueba de
   round-trip con posición negativa.
5. **El validador del SDK acepta `w:cr` y `w:lastRenderedPageBreak` dentro de un run.** Se fijó con
   una aserción de esquema sobre el paquete producido a partir de un archivo ajeno con ese contenido,
   y esos nodos se preservan aunque no se modelen.
6. **xUnit no puede transportar suplentes sueltos en `[InlineData]`.** Las cadenas con un suplente
   suelto se serializan con el carácter de reemplazo, de modo que casos distintos colisionaban
   («Duplicate test case») y llegaba un texto válido a la prueba. Se cambió la teoría por hechos con
   las cadenas literales.

### Qué **no** se ha implementado (y por tanto no se afirma)

- Las tabulaciones heredadas (estilo, numeración e intervalo por defecto): `Paragraph.TabStops` solo
  informa de las que define el propio párrafo.
- Las alineaciones `start`, `end` y `num` (lectura con `UnsupportedOfficeFeatureException`, marcado
  preservado) y el atributo `w:clear` de `w:br` (preservado, no reportado).
- El contenido no modelado de un run (dibujos, campos, `w:noBreakHyphen`, tabulaciones absolutas…) no
  aparece en `Run.Content`, aunque se preserva.
- Estilos, listas, tablas editables, secciones, encabezados/pies, imágenes, hipervínculos y
  búsqueda/reemplazo: sin cambios respecto a 2.1.
- La decisión de diseño (tipo de valor, signo de la posición, escritura canónica, validación de texto)
  está en `docs/adr/0010-tab-stops-and-breaks.md`.

## Incremento 2.3 — Referencias de estilo y catálogo de estilos (solo lectura)

### Alcance

Acotado expresamente a **referenciar** estilos y **leer** el catálogo, sin crear ni editar
definiciones. Incluye:

- `Paragraph.StyleId` (`w:pPr/w:pStyle/@w:val`) y `Run.StyleId` (`w:rPr/w:rStyle/@w:val`), ambos
  `string?` y tri-state (`null` = heredar), validados antes de escribir y nunca resueltos.
- `WordDocument.Styles` (catálogo en orden de documento; vacío si no hay parte de estilos) y
  `WordDocument.FindStyle(string)` (comparación ordinal, primera coincidencia).
- `WordStyle` con `Id`, `Name`, `Type` (`WordStyleType`), `IsDefault`, `BasedOnId`, `OwnFormatting` y
  `EffectiveFormatting`.
- `WordStyleFormatting`: estructura de solo lectura con doce miembros tri-state (fuente, tamaño,
  negrita, cursiva, color; alineación, sangrías izquierda/derecha/primera línea/francesa y espaciado
  anterior/posterior) y resolución de la herencia `w:basedOn`.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Debug` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Debug` | **306 superadas, 0 con error, 0 omitidas** (306 total). |
| `dotnet test OfficeEngine.sln -c Release` | **306 superadas, 0 con error, 0 omitidas** (306 total). |

Salida textual del resumen de `dotnet test` (las duraciones varían entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   306, Omitido:     0, Total:   306
```

El informe TRX de Release está en `artifacts/testresults/phase2/inc23/coverage/phase2_inc23_release.trx`
y la cobertura en `artifacts/testresults/phase2/inc23/coverage/**/coverage.cobertura.xml`.

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Word/WordStyle.cs` | 240 / 244 | 98,4 % | 98,5 % |
| `Word/WordStyleFormatting.cs` | 108 / 108 | 100 % | 92,8 % |
| `Word/StyleReference.cs` | 20 / 20 | 100 % | 100 % |
| `Word/PlainText.cs` | 70 / 70 | 100 % | 100 % |
| `Word/Paragraph.cs` | 662 / 662 | 100 % | 94,6 % |
| `Word/Run.cs` | 512 / 516 | 99,2 % | 91,7 % |
| **Total de la solución** | 1405 / 1489 | 94,3 % | 88,5 % |

`WordStyle.cs` deja sin cubrir únicamente las dos líneas del `throw` defensivo de `ToWordStyleType`
para un valor de `ST_Style` distinto de los cuatro modelados: la rama es inalcanzable (el propio SDK
solo conoce esos cuatro valores) y se mantiene por coherencia con la rama análoga de `ReadBreakType`
en `Run.cs`. `WordDocument.cs` (228 / 266) incluye rutas de la Fase 1 no tocadas por este incremento;
los miembros nuevos (`Styles`, `FindStyle`, la lectura de la parte de estilos) sí están cubiertos.

### Pruebas nuevas del incremento (99 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordStyleReferenceTests.cs` (43) | Referencias de párrafo y de run: valor por defecto, lectura, referencia colgante, presencia sin `w:val`, escritura e inserción en orden de esquema, reemplazo, borrado (con poda de `w:pPr`/`w:rPr` vacío) sin perder el resto del formato, rechazo de vacío/solo espacios/tabulador/salto/demasiado larga/carácter no representable sin dejar marcado, límite de 253, carácter astral, espacio interior, round-trip por stream y por archivo, coexistencia de ambas referencias y validación de esquema. |
| `Word/WordStyleCatalogTests.cs` (49) | Catálogo: vacío sin parte de estilos y sin raíz, orden de documento, identidad (id/nombre/tipo/`w:default`/`basedOn`), los cuatro `ST_Style`, `FindStyle` (ordinal, primera coincidencia, argumentos), formato propio de carácter y de párrafo (incluidas negrita/cursiva desactivadas, los cinco alineamientos, fuente solo `hAnsi`, omisión del interlineado, color de tema y color sin valor), los cuatro mapeos de enumeración mal formados, alineación no modelada, herencia `w:basedOn` (una y tres niveles, anulación, `w:cs`, color de tema, base colgante), ciclo y autorreferencia, igualdad de valor de `WordStyleFormatting` y `ToString`. |
| `Word/WordStyleInteropTests.cs` (7) | Interoperabilidad: validación de esquema del paquete producido con referencias, lectura con el SDK de `w:pStyle`/`w:rStyle`, ausencia de una parte de estilos inventada, lectura de un archivo ajeno con `w:styles` comparando con el marcado, conservación byte a byte de `styles.xml` al editar el cuerpo, round-trip de una referencia colgante y round-trip por archivo. |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes producidos con referencias de
   párrafo y de run, de los documentos con `w:styles` leídos y editados y del round-trip de una
   referencia colgante.
2. **Lectura con el SDK** de `w:pStyle`/`w:rStyle` y del catálogo
   (`MainDocumentPart.StyleDefinitionsPart.Styles`).
3. **Lectura cruda** de `word/styles.xml` para comprobar que editar el cuerpo no altera la parte de
   estilos byte a byte.
4. **Fixtures creados con el SDK** (definiciones `w:style` y XML crudo) leídos por OfficeEngine:
   identidad, herencia `w:basedOn` en cadena, ciclo, base colgante, valor de enumeración mal formado y
   color de tema.

### Hallazgos durante las pruebas y acciones tomadas

1. **El validador del SDK no comprueba integridad referencial.** Un `w:pStyle`/`w:rStyle` que nombra
   un estilo inexistente —incluso en un documento sin `styles.xml`— valida con **0 errores**, y Word
   también ignora la referencia colgante (comprobado con la sonda). Se decidió que una referencia
   colgante **no** es un error: se guarda y se lee tal cual. Habría sido fácil «endurecer» y lanzar,
   pero eso habría inventado una regla que el formato no impone.
2. **`w:styleId` y `w:type` ausentes son válidos** (0 errores) y se modelan como `null`; en cambio
   `w:name`, `w:basedOn`, `w:pStyle` y `w:rStyle` presentes **sin** `w:val` son error de esquema y se
   reportan como `InvalidOfficeFileException`.
3. **Los fixtures «realistas» hay que construirlos con XML crudo.** Asignar un valor de enumeración
   inválido a una propiedad tipada del SDK **valida inmediatamente**, así que un `w:type`, un `w:b` o
   un `w:default` mal formados se crean con `SetAttribute` para ejercitar de verdad la ruta
   `FormatException → InvalidOfficeFileException` (mismo hallazgo que en 2.2b).
4. **El color de un estilo: manda la presencia del elemento.** Un `w:color` en el estilo derivado
   —aunque sea `auto` o de tema, sin RGB fijo— sustituye el valor heredado y **no** filtra el RGB del
   estilo base. Se probó explícitamente para que la herencia no cuele un color que el derivado quiso
   anular.
5. **Un `w:rFonts` que solo define la fuente de script complejo (`w:cs`) no cambia la fuente
   occidental** y, por tanto, hereda. Se fijó con una prueba dedicada.
6. **La rama defensiva de `ToWordStyleType` es inalcanzable.** `ST_Style` define exactamente cuatro
   valores, todos modelados, así que las dos líneas del `throw` quedan sin cubrir (igual que la rama
   análoga de `ReadBreakType` en `Run.cs`); se mantiene por coherencia con el resto de mapeos.

### Qué **no** se ha implementado (y por tanto no se afirma)

- Crear, editar o eliminar definiciones de estilo: el catálogo es de solo lectura y `styles.xml` se
  conserva tal cual.
- Resolver `w:docDefaults`, el estilo por defecto del tipo, los estilos vinculados (`w:link`) ni el
  estilo siguiente (`w:next`): `EffectiveFormatting` resuelve solo la cadena `w:basedOn`.
- El formato de estilo completo (resaltado, tachado, bordes, `w:tabs`, interlineado…):
  `OwnFormatting`/`EffectiveFormatting` cubren fuente, tamaño, negrita, cursiva, color, alineación,
  sangrías y espaciado anterior/posterior.
- Aplicar a un párrafo o run el formato efectivo de un estilo: la API lee referencias y catálogo, no
  compone el resultado final que Word renderizaría.
- La decisión de diseño (tri-state, validación, ausencia de integridad referencial, fusión por
  presencia del color, exclusión del interlineado) está en
  `docs/adr/0011-style-references-and-catalogue.md`.

## Incremento 2.4 — Listas numeradas y con viñetas (referencias, catálogo y creación)

### Alcance

Acotado expresamente a **referenciar** listas, **leer** el catálogo y **crear** listas de un nivel,
sin editar definiciones existentes. Incluye:

- `Paragraph.ListId` (`w:pPr/w:numPr/w:numId`) y `Paragraph.ListLevel`
  (`w:pPr/w:numPr/w:ilvl`), ambos `int?` y tri-state (`null` = no se nombra; `0` en `ListId` es el
  valor reservado «no numerado»), validados antes de escribir y nunca resueltos.
- `Paragraph.ApplyList(WordList, int level = 0)`, que sitúa un párrafo en una lista y un nivel.
- `WordDocument.Lists` (catálogo en orden de documento; vacío si no hay parte de numeración ni raíz)
  y `WordDocument.FindList(int)` (primera coincidencia con `w:numId`; omite una definición cuyo
  identificador no se pueda leer).
- `WordList` (`Id`, `AbstractNumberId`, `Levels`, `FindLevel`, `ToString`) y `WordListLevel`
  (`Level`, `Start`, `Format`, `Text`, `Justification`).
- `WordDocument.AddList(WordListFormat)`: crea `word/numbering.xml` si hace falta, asignando
  identificadores y añadiendo un `w:abstractNum` (antes del primer `w:num`) y un `w:num` (al final)
  con un solo nivel, numerado o con viñeta.
- `WordListFormat` (`Bullet`, `Decimal`, `LowerLetter`, `UpperLetter`, `LowerRoman`, `UpperRoman`) y
  `WordListJustification` (`Left`, `Center`, `Right`).
- Inserción en orden de esquema dentro de `w:numPr` (`w:ilvl` antes de `w:numId`), preservando el
  marcado no modelado del contenedor.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Debug` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Debug` | **388 superadas, 0 con error, 0 omitidas** (388 total). |
| `dotnet test OfficeEngine.sln -c Release` | **388 superadas, 0 con error, 0 omitidas** (388 total). |

Salida textual del resumen de `dotnet test` (las duraciones varían entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   388, Omitido:     0, Total:   388
```

El informe TRX de Release está en `artifacts/testresults/phase2/inc24/coverage/phase2_inc24_release.trx`
y la cobertura en `artifacts/testresults/phase2/inc24/coverage/**/coverage.cobertura.xml`.

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Word/WordList.cs` | 84 / 86 | 97,7 % | 96,4 % |
| `Word/WordListLevel.cs` | 136 / 140 | 97,1 % | 95,5 % |
| `Word/ListReference.cs` | 32 / 32 | 100 % | 90,0 % |
| `Word/Paragraph.cs` | 776 / 776 | 100 % | 94,9 % |
| `Word/WordDocument.cs` | 370 / 408 | 90,7 % | 84,4 % |
| **Total de la solución** | 1659 / 1746 | 95,0 % | 89,5 % |

`WordList.cs` deja sin cubrir la única línea del `return null` de `FindAbstract` para una raíz
ausente: cuando `Lists` devuelve elementos la raíz siempre existe, así que la rama es defensiva.
`WordListLevel.cs` deja sin cubrir las dos líneas del `throw` defensivo de `ToWordListJustification`
para un valor de `ST_LevelJustification` distinto de los tres modelados: la rama es inalcanzable (el
SDK solo conoce esos tres y el validador rechaza el resto al leer). Ambas se mantienen por coherencia
con el resto de mapeos. `WordDocument.cs` incluye rutas de la Fase 1 y de `TryDelete` no tocadas por
este incremento; los miembros nuevos (`Lists`, `FindList`, `AddList`, sus ayudantes y el mapeo de
formatos) sí están cubiertos.

### Pruebas nuevas del incremento (82 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordListReferenceTests.cs` (34) | Referencias de lista: valor por defecto, lectura, `w:numId="0"`, presencia sin `w:val`, valor mal formado, `w:ilvl` fuera del rango que la API escribe, escritura y inserción en orden de esquema (`w:numPr` en `w:pPr`, `w:ilvl` antes de `w:numId`), reemplazo, borrado (con poda de `w:numPr`/`w:pPr` vacío) sin perder el resto del formato ni el otro miembro, rechazo de valores fuera de rango sin dejar marcado, preservación de propiedades de párrafo no modeladas, `ApplyList` (nivel por defecto, nivel indicado, `null` y rango), round-trip por stream y por archivo. |
| `Word/WordListCatalogTests.cs` (39) | Catálogo: vacío sin parte de numeración y sin raíz, orden de documento, identidad (`Id`/`AbstractNumberId`/`ToString`), niveles leídos de la definición abstracta, referencia colgante (sin niveles), identidad del nivel (`Level`/`Start`/`Format`/`Text`/`Justification`), `FindLevel` (coincidencia, ausencia, rango), los seis formatos modelados, formato no modelado (`chicago`), formato y justificación mal formados o sin `w:val`, ausencia de formato/inicio/texto/justificación, justificación no legible (`both`), `w:ilvl` ausente, `w:numId` ausente o mal formado, `w:abstractNumId` ausente o sin `w:val`, `FindList` que omite un identificador ilegible, y `AddList` (crea la parte, un nivel en orden de esquema, viñeta con fuente `Symbol`, todos los formatos, orden `abstractNum`-antes-de-`num` —incluidas las instancias preexistentes—, identificadores incrementales, rechazo de formato no definido sin tocar el documento, round-trip). |
| `Word/WordListInteropTests.cs` (9) | Interoperabilidad: validación de esquema del paquete con listas, lectura con el SDK de `w:numId`/`w:ilvl`, ausencia de una parte de numeración inventada cuando solo hay una referencia colgante, lectura de un archivo ajeno con niveles comparando con el marcado, conservación byte a byte de `numbering.xml` al editar el cuerpo, preservación de una definición ajena (con `w:nsid` y nivel no modelado) al añadir una lista, edición del formato del párrafo sin romper la lista, round-trip de una referencia colgante y round-trip por archivo. |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con listas creadas y aplicadas, de
   los documentos con `numbering.xml` leídos y editados y del round-trip de una referencia colgante.
2. **Lectura con el SDK** de `w:numPr` (`w:numId`/`w:ilvl`), de la parte de numeración
   (`MainDocumentPart.NumberingDefinitionsPart.Numbering`) y del orden de sus hijos.
3. **Lectura cruda** de `word/numbering.xml` para comprobar que editar el cuerpo no altera la parte de
   numeración byte a byte y que añadir una lista no toca las definiciones existentes.
4. **Fixtures creados con el SDK** (definiciones `w:abstractNum`/`w:num` y XML crudo, incluidas
   cabeceras sin `w:val` y valores mal formados) leídos por OfficeEngine: identidad, niveles, formato
   no modelado, justificación no válida y referencias colgantes.

### Hallazgos durante las pruebas y acciones tomadas

1. **El validador del SDK no comprueba integridad referencial en la numeración.** Un `w:numId` que
   nombra una lista inexistente —incluso en un documento sin `numbering.xml`—, `w:numId="0"`,
   `w:ilvl="12"` y un `w:num` que apunta a un `w:abstractNum` ausente validan con **0 errores**, y
   Word los ignora. Se decidió que ninguna de esas situaciones es un error: se guardan y se leen tal
   cual (`Levels` queda vacío cuando falta la abstracción).
2. **En cambio, hay miembros que el esquema sí exige.** `w:num` sin `w:numId`, `w:num` sin
   `w:abstractNumId` (o con el elemento sin `w:val`), `w:abstractNum` sin `w:abstractNumId` y `w:lvl`
   sin `w:ilvl` son errores de validación; se reportan como `InvalidOfficeFileException` y por eso
   `WordList.Id`/`AbstractNumberId`/`WordListLevel.Level` no son nullable (a diferencia de
   `w:styleId`, que el validador acepta ausente).
3. **`ST_LevelJustification` tiene exactamente tres valores.** A diferencia de la alineación de
   párrafo (`ST_Jc`), el validador rechaza `both` y `start`, así que un `w:lvlJc` no legible es
   `InvalidOfficeFileException`, no `UnsupportedOfficeFeatureException`. Solo `ST_NumberFormat`
   (unos 63 valores) da el caso «válido pero no modelado» → `UnsupportedOfficeFeatureException`.
4. **El orden de la parte importa.** `CT_Numbering` exige que **todos** los `w:abstractNum` (y
   `w:numPicBullet`) precedan a **todos** los `w:num`: `AddList` inserta la abstracción antes de la
   primera instancia, no la añade al final. Se probó tanto con la parte vacía como con una
   definición preexistente.
5. **El SDK inserta las propiedades tipadas en orden de esquema.** Un `w:lvl` construido con
   `StartNumberingValue`, `NumberingFormat`, `LevelText`, `LevelJustification`,
   `PreviousParagraphProperties` y `NumberingSymbolRunProperties` se serializa en orden válido sin
   ordenar los hijos a mano (comprobado con la sonda y con la validación).
6. **La viñeta se escribe como Word.** El nivel de viñeta usa `w:lvlText` `U+F0B7` y
   `w:rPr/w:rFonts` `Symbol` (no el carácter `U+2022`), que es lo que produce la lista con viñetas
   predeterminada de Word; se fijó con una prueba.
7. **La rama defensiva de `ToWordListJustification` es inalcanzable**, igual que la de
   `ToWordStyleType` en 2.3: `ST_LevelJustification` solo define tres valores, todos modelados, y el
   resto no se puede leer. Se mantiene por coherencia con los demás mapeos.
8. **Rechazar antes de mutar.** `ListId`/`ListLevel` validan el rango antes de tocar el árbol y
   `ApplyList` lee `list.Id` antes de escribir, de modo que un valor rechazado (o una lista ilegible)
   no deja marcado parcial; se probó explícitamente.

### Qué **no** se ha implementado (y por tanto no se afirma)

- Crear listas de **más de un nivel**: `AddList` produce un solo nivel (0) y no se añaden niveles,
  `w:startOverride`, reinicios, `w:numPicBullet` ni numeración ligada a estilos.
- Editar o eliminar definiciones de lista existentes: el catálogo es de solo lectura y
  `numbering.xml` se conserva tal cual salvo las listas que `AddList` añade.
- Exponer las propiedades de párrafo/carácter de un nivel (`w:pPr`/`w:rPr` del nivel): se preservan,
  pero no se modelan (la sangría que el nivel aplica no se lee).
- Componer el resultado final que Word renderizaría: la API lee referencias y catálogo y aplica una
  referencia, no resuelve la numeración efectiva de un párrafo.
- La decisión de diseño (tri-state, validación, ausencia de integridad referencial, subconjunto de
  formatos, orden de esquema, viñeta `Symbol`) está en
  `docs/adr/0012-lists-references-catalogue-and-creation.md`.

## Incremento 2.5a — Tablas: estructura y contenido

### Alcance

Acotado expresamente a la **estructura** y el **contenido** de una tabla, sin modelar todavía su
formato. El incremento 2.5 completo se dividió en 2.5a/2.5b/2.5c (como el 2.2/2.2b) **en lugar de
declarar hecho algo a medias**. Incluye:

- `WordDocument.Tables` (tablas directas del cuerpo, en orden de documento) y
  `WordDocument.AddTable(filas, columnas)`, que añade la tabla al final del cuerpo.
- `Table` (`Rows`, `RowCount`, `ColumnCount`, `AddRow`, `ToString`), `TableRow` (`Cells`, `CellCount`,
  `ToString`) y `TableCell` (`Paragraphs`, `Text`, `AddParagraph`, `ToString`).
- Creación en orden de esquema (`w:tblPr` vacío, `w:tblGrid` con una `w:gridCol` por columna y filas
  con una `w:tc` por columna, cada una con un párrafo vacío), con un ancho de columna predeterminado
  (`2400` twips) porque el ancho no se modela todavía.
- `WordDocument.Text` pasa a incluir el texto de las celdas: una línea por fila, celdas separadas por
  tabulador, y una línea por párrafo dentro de una celda.
- El marcado de tabla no modelado (ancho, bordes, sombreado, `w:gridSpan`/`w:vMerge`, tablas
  anidadas) se **preserva** al editar: es el criterio «tablas existentes no se destruyen al editar».

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Debug` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Debug` | **423 superadas, 0 con error, 0 omitidas** (423 total). |
| `dotnet test OfficeEngine.sln -c Release` | **423 superadas, 0 con error, 0 omitidas** (423 total). |

Salida textual del resumen de `dotnet test` (las duraciones varían entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   423, Omitido:     0, Total:   423
```

El informe TRX de Release está en `artifacts/testresults/phase2/inc25/coverage/phase2_inc25_release.trx`
y la cobertura en `artifacts/testresults/phase2/inc25/coverage/**/coverage.cobertura.xml`.

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Word/Table.cs` | 42 / 42 | 100 % | 100 % |
| `Word/TableRow.cs` | 8 / 8 | 100 % | 100 % |
| `Word/TableCell.cs` | 36 / 36 | 100 % | 100 % |
| `Word/Paragraph.cs` | 776 / 776 | 100 % | 94,9 % |
| `Word/WordDocument.cs` | 412 / 450 | 91,5 % | 86,5 % |
| **Total de la solución** | 1723 / 1810 | 95,2 % | 89,7 % |

`WordDocument.cs` deja sin cubrir rutas defensivas de la Fase 1 (los `catch` de
`WordDocument.Create`/`Save` y el cuerpo de `TryDelete`); los miembros nuevos (`Tables`, `AddTable` y
el recorrido de bloques de `Text`) están cubiertos al 100 %.

### Pruebas nuevas del incremento (35 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordTableTests.cs` (26) | Modelo: cuerpo sin tablas; `AddTable` (filas/columnas, celda con un párrafo vacío, posición al final del cuerpo, dimensiones no positivas, documento liberado); orden de las tablas; `ColumnCount` sin rejilla; tabla sin filas; `CellCount`; `AddRow` (celdas por columna de la rejilla, también en una tabla ajena); `TableCell.Text` (concatenación de párrafos, reemplazo que conserva `w:tcPr` y el marcado no modelado, valor vacío, rechazo de tabuladores/saltos y de `null`); `AddParagraph`; edición de un párrafo de la celda por su propia API; `Text` (incluye celdas, separa con tabulador y salto de línea, una línea por párrafo, no incluye tablas anidadas ni contenido no modelado); `ToString`. |
| `Word/WordTableInteropTests.cs` (8) | Interoperabilidad: validación de esquema; orden de los hijos de `w:tbl`; una `w:gridCol` por columna; lectura del texto con el SDK; intercalado de párrafos y tablas; texto persistente entre guardados; preservación del marcado no modelado (ancho, sombreado, `w:gridSpan`) al editar el texto; preservación de una tabla anidada al reescribir el contenido de la celda; round-trip. |
| `Word/WordInteropTests.cs` (+1 neto) | Actualizado: una tabla modelada ya no genera diagnóstico (`OpenedDocument_WithTable_ExposesItAndKeepsItWhenTheBodyIsEdited`, `...WithModelledContent_HasNoDiagnostics`); el diagnóstico de contenido no modelado se conserva con una `w:sdt` (`OpenedDocument_WithUnmodelledBodyContent_ReportsDiagnosticAndPreservesIt`). |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con tablas creadas, con texto y
   con una fila añadida, y de los documentos ajenos editados.
2. **Lectura con el SDK** de los hijos de `w:tbl` y de su orden, de la rejilla (`w:gridCol/@w:w`) y
   del texto de las celdas, así como del orden de los hijos del cuerpo (`w:p`/`w:tbl`).
3. **Lectura cruda** de `word/document.xml` para comprobar que editar el texto de una celda conserva
   `w:tcPr`, `w:gridSpan`, el sombreado y `w:tblW`, y que una tabla anidada sobrevive.
4. **Fixtures creados con el SDK** (tablas con y sin rejilla, con formato no modelado, anidadas y
   cuerpos con una `w:sdt`) leídos por OfficeEngine.

### Hallazgos durante las pruebas y acciones tomadas

1. **El SDK no reordena los hijos de `w:tblPr`/`w:tcPr`.** A diferencia de `w:lvl` en 2.4, construir
   una celda con el ancho, la combinación y el sombreado en un orden distinto del que exige el
   esquema produce XML inválido (lo detectó el validador). Se confirma la decisión de insertar siempre
   en orden de esquema; los hijos que se escriben hoy (creación de la tabla) ya están en orden.
2. **El esquema permite cosas que no sirven.** Un `w:tblPr` vacío, una `w:tc` sin `w:p`, una
   `w:gridCol` sin `w:w` y una tabla sin filas o sin columnas validan con **0 errores**; aun así
   `AddTable` exige una fila y una columna como mínimo, y cada celda recibe un párrafo, que es lo que
   Word necesita.
3. **Pérdida de datos evitada en `TableCell.Text`.** La primera versión del *setter* borraba todo
   salvo `w:tcPr`, lo que habría eliminado en silencio una **tabla anidada** que la versión no modela.
   Se cambió para eliminar **solo los párrafos** y conservar cualquier otro contenido de bloque; hay
   prueba con una tabla anidada.
4. **`WordDocument.Text` se amplió, no se rompió.** Al modelar las tablas, dejar su texto fuera de
   `Text` habría sido una omisión silenciosa; ahora una tabla aporta una línea por fila con las celdas
   separadas por tabulador. Es un cambio documentado de contrato (ADR 0013) y se actualizaron las
   pruebas y el `README`.
5. **`w:tblCellMar` no acepta `w:start`/`w:end`.** La sonda mostró que el validador los rechaza (usa
   `w:left`/`w:right`). Anotado para cuando se modelen los **márgenes de celda**, que quedaron **fuera
   del 2.5b** (ver más abajo) y siguen preservándose.
6. **Un cuerpo que termina en tabla es válido**, así que `AddTable` puede ser la última operación
   sobre el cuerpo sin necesidad de añadir un párrafo detrás.

### Qué **no** se ha implementado (y por tanto no se afirma)

- **Propiedades de tabla y de celda**: ancho (de tabla y de celda), alineación, disposición
  (`w:tblLayout`), bordes, sombreado y alineación vertical. Se preservan, pero no se modelan
  (incremento 2.5b).
- **Combinación de celdas** (`w:gridSpan`/`w:vMerge`): se preserva, pero no se modela en este
  incremento; la modela el 2.5c (ver su sección).
- **Tablas anidadas** dentro de una celda: se preservan, pero no las devuelve `Tables` ni
  `TableCell.Paragraphs`, y su texto no entra en `WordDocument.Text`.
- **Añadir o quitar columnas**: las columnas se fijan al crear la tabla; solo se añaden filas
  (`Table.AddRow`).
- **Eliminar o mover** tablas, filas o celdas: solo se añaden al final, igual que los párrafos.
- **Medir la página**: el ancho de columna de una tabla creada es un valor predeterminado fijo
  (2400 twips), no el resultado de medir la sección.
- La decisión de diseño (tabla como bloque, `w:tblPr` vacío, ancho predeterminado, escritura de texto
  que no descarta contenido, `Text` con celdas) está en
  `docs/adr/0013-tables-structure-and-content.md`.

## Incremento 2.5b — Tablas: propiedades de tabla y de celda

### Alcance

Modela el formato de una tabla que 2.5a **preservaba** pero no leía ni escribía. Incluye:

- **Tabla** (`Table`): `Width` (ancho preferido, `w:tblW`), `Alignment` (`w:jc`), `Layout`
  (`w:tblLayout`), `ShadingColor` (`w:shd`) y seis bordes (`w:tblBorders`):
  `BorderTop`, `BorderLeft`, `BorderBottom`, `BorderRight`, `BorderInsideHorizontal`,
  `BorderInsideVertical`.
- **Celda** (`TableCell`): `Width` (`w:tcW`), `VerticalAlignment` (`w:vAlign`), `ShadingColor`
  (`w:shd`) y sus cuatro bordes exteriores (`w:tcBorders`).
- **Tipos de valor**: `TableWidth` (`Type` + `Value`, con `Points`/`Percentage`/`Auto`),
  `TableBorder` (`Style` + `SizePoints` + `Color`) y las enumeraciones `TableWidthType`,
  `TableAlignment`, `TableLayout`, `TableVerticalAlignment` y `BorderStyle` (los 27 valores no
  decorativos de `ST_Border`).
- **Semántica tri-state** en todas las propiedades: `null` significa «el elemento no está» (lo
  hereda); asignar `null` lo elimina.
- **Error tipado antes que invención**: un valor que OOXML define pero la versión no modela (unidad
  `nil`, patrón de sombreado distinto de `clear`, borde decorativo) lanza
  `UnsupportedOfficeFeatureException` y conserva el marcado; un valor mal formado o un atributo
  obligatorio ausente lanza `InvalidOfficeFileException`.
- **Escritura en orden de esquema** (ayudante interno `SchemaOrder`) y preservación de lo no
  modelado (aristas diagonales, atributos no modelados, estilo, `w:tblLook`).

**Fuera del 2.5b** (y por tanto no afirmado): los **márgenes de celda** (`w:tblCellMar`/`w:tcMar`) se
siguen preservando sin modelar, y la **combinación de celdas** (`w:gridSpan`/`w:vMerge`) quedó fuera
de este incremento; la modela el 2.5c (ver su sección).

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release` | **472 superadas, 0 con error, 0 omitidas** (472 total). |

Salida textual del resumen de `dotnet test` (las duraciones varían entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   472, Omitido:     0, Total:   472
```

El informe TRX de Release está en
`artifacts/testresults/phase2/inc25b/coverage/phase2_inc25b_release.trx` y la cobertura en
`artifacts/testresults/phase2/inc25b/coverage/**/coverage.cobertura.xml`.

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Word/Table.cs` | 122 / 122 | 100 % | 93,2 % |
| `Word/TableCell.cs` | 104 / 104 | 100 % | 96,3 % |
| `Word/TableBorder.cs` | 85 / 85 | 100 % | 92,1 % |
| `Word/TableWidth.cs` | 80 / 80 | 100 % | 98,1 % |
| `Word/TableShading.cs` | 17 / 17 | 100 % | 93,8 % |
| `Word/BorderEdges.cs` | 16 / 16 | 100 % | 91,7 % |
| `Word/SchemaOrder.cs` | 10 / 11 | 90,9 % | 75 % |
| `Word/TableRow.cs` | 4 / 4 | 100 % | 100 % |
| **Total de la solución** | 2118 / 2206 | 96,0 % | 90,4 % |

La única línea sin cubrir es el `element.Remove()` de `SchemaOrder.Insert`, una salvaguarda para el
caso de insertar un elemento que ya está enlazado: los escritores de este incremento siempre insertan
elementos recién construidos, así que hoy no es alcanzable desde la API.

### Pruebas nuevas del incremento (49 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordTableFormattingTests.cs` (44) | Modelo y errores: una tabla nueva no tiene formato; ancho de tabla en puntos (twips), porcentaje y `auto`, sin `w:type` (se asume `dxa`) y sin medida (0), unidad `nil` no modelada y medida mal formada; alineación (los tres valores, elemento, valor no modelado, valor no definido) y disposición (`fixed`/`autofit`, sin `w:type`, valor no modelado); sombreado (relleno sólido `clear`, `auto`/temático, patrón no modelado, `w:val` ausente y color mal formado); bordes de tabla (las seis aristas, borrado, escritura en octavos de punto, medio punto por defecto, los 27 estilos, estilo decorativo, `w:val` ausente y `w:sz` mal formado, color mal formado); validación de `TableWidth`/`TableBorder` (no finito, negativo, demasiado grande) y de los enumerados; igualdad y `ToString`; orden de esquema de `w:tblPr` en torno a marcado no modelado; tabla sin `w:tblPr` (lee vacío y lo crea al escribir); round-trip por stream; celda: ancho y alineación vertical, los tres valores, las cuatro aristas, nunca toca `insideH` ni la diagonal, poda de `w:tcPr` vacío, orden de esquema de `w:tcPr` y limpieza inerte cuando no hay nada que limpiar. |
| `Word/WordTableFormattingInteropTests.cs` (5) | Interoperabilidad con el SDK: paquete válido con formato de tabla y de celda y lectura de `w:tblW`/`w:jc`/`w:tblLayout`/`w:tblBorders`/`w:tcW`/`w:vAlign`/`w:shd`; orden de los hijos de `w:tblPr` y de `w:tcPr`; los 27 estilos de borde producen documentos válidos; editar el formato de una tabla ajena conserva su estilo, sus márgenes de celda, su `w:tblLook` y su borde interno, y el paquete sigue validando. |
| `Support/OoxmlValidation.cs` (+4 ayudantes) | Fixtures de tabla/celda con propiedades exactas y lectores del `w:tblPr`/`w:tcPr` del primer elemento, construidos con el SDK (no con la librería). |

Las 423 pruebas del 2.5a **siguen pasando** sin cambios: no hay regresiones.

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con formato de tabla y de celda,
   de los 27 estilos de borde y de los documentos ajenos editados.
2. **Lectura con el SDK** de los valores escritos (`w:tblW`, `w:jc`, `w:tblLayout`, `w:tblBorders`,
   `w:shd`, `w:tcW`, `w:tcBorders`, `w:vAlign`) y del orden de los hijos de `w:tblPr`/`w:tcPr`.
3. **Fixtures creados con el SDK** (anchos con y sin `w:type`, sombreados con patrón no modelado,
   bordes con `w:val` ausente o decorativo, color mal formado, `w:jc` con valor no modelado, tablas
   sin `w:tblPr`) leídos por OfficeEngine, que debe leer el valor, fallar con error tipado o crear lo
   que falte.

### Hallazgos durante las pruebas y acciones tomadas

1. **`BorderValues` es un `readonly struct`, no un `enum`.** `Enum.Parse` no sirve para mapear el
   estilo; se construye desde la cadena del XML. Se evita duplicar la tabla de estilos leyéndola al
   revés: una única tabla `string → BorderStyle` y su inversa, de modo que escribir y leer no pueden
   divergir (lo fija una prueba que recorre los 27 estilos).
2. **Un `w:type` omitido no es «no soportado».** `ST_TblWidth` y `ST_TblLayoutType` definen el valor
   por defecto (`dxa` y `autofit`), así que tratarlos como no soportados volvería ilegible un archivo
   válido. Se leen con su valor por defecto; los atributos obligatorios (`w:val`) sí fallan.
3. **Una tabla sin `w:tblPr` es ilegible para la escritura, no para la lectura.** `CT_Tbl` exige
   `w:tblPr` como primer hijo; si un paquete mal formado no lo tiene, leer devuelve «sin
   propiedades» y la primera escritura **crea** el elemento en la posición 0. Hay prueba de que el
   documento resultante valida.
4. **`w:tcPr` no se puede podar a ciegas.** Un `w:tcPr` sin hijos puede seguir teniendo atributos
   (por ejemplo `w:tcPrChange` con atributos): la poda comprueba hijos **y** atributos antes de
   eliminar el elemento.
5. **La celda no posee `insideH`/`insideV`.** Escribir un borde de celda solo toca su arista; el
   borde interno de una tabla ajena y la diagonal de la celda sobreviven a cualquier escritura (hay
   prueba que las cuenta y las compara).
6. **`w:jc` de `w:tblPr` usa `TableRowAlignmentValues`.** La sonda confirmó que el SDK lo valida con
   esa enumeración (no con `JustificationValues`) y que `start`/`end` no son válidos ahí; el mapeo
   usa `left`/`center`/`right` y cualquier otro valor se reporta como no legible.

### Qué **no** se ha implementado (y por tanto no se afirma)

- **Márgenes de celda** (`w:tblCellMar`/`w:tcMar`): se preservan, pero no se modelan.
- **Combinación de celdas** (`w:gridSpan`/`w:vMerge`): se preserva, pero no se modela en este
  incremento; la modela el 2.5c (ver su sección).
- **Bordes decorativos** (`ST_Border` con valores `apples`, `balloons`…): se preservan; leerlos falla
  con `UnsupportedOfficeFeatureException`.
- **Patrones de sombreado** distintos de `clear` (tramas, porcentajes): se preservan; leerlos falla
  con `UnsupportedOfficeFeatureException` (asignar un color escribe siempre un relleno sólido).
- **La rejilla** (`w:tblGrid`): fijar el ancho preferido de la tabla no la reescribe; las columnas
  siguen fijándose al crear la tabla.
- **Ancho de columna por columna**: se modela el ancho **preferido** de la tabla y el de una celda,
  no `w:tblGrid/w:gridCol` de cada columna.
- La decisión de diseño (una propiedad por arista, tri-state, error tipado, orden de esquema,
  `w:tblPr` que nunca se elimina) está en
  `docs/adr/0014-table-and-cell-properties.md`.

## Incremento 2.5c — Tablas: combinación de celdas

### Alcance

Modela la combinación de celdas, que 2.5a y 2.5b **preservaban** pero no leían ni escribían, y añade
las operaciones para combinar y dividir. Incluye:

- **Lectura/escritura de las marcas**: `TableCell.ColumnSpan` (`w:gridSpan`, un `int` con 1 por
  defecto: una columna y «sin span» son lo mismo) y `TableCell.VerticalMerge`
  (`TableVerticalMerge?` sobre `w:vMerge`: `null` = no participa, `Restart` = `w:val="restart"`,
  `Continue` = `w:val="continue"` o `w:val` ausente).
- **Operaciones horizontales** (`TableRow`): `MergeCells(startIndex, count)` y
  `UnmergeCell(cellIndex)`.
- **Operaciones verticales** (`Table`): `MergeCellsVertically(firstRowIndex, cellIndex, rowCount)` y
  `UnmergeCellsVertically(...)`.
- **Combinar no destruye contenido**: el combinado horizontal **mueve** a la celda que absorbe todos
  los hijos (menos `w:tcPr`) de las celdas que elimina, así que un párrafo o una **tabla anidada**
  viajan con el contenido; el combinado vertical **solo escribe las marcas** y no toca el contenido
  de las celdas (lo decide el llamante).
- **Ancho de la celda combinada**: es la suma de los anchos cuando todos son del mismo tipo; si falta
  alguno, son de tipos distintos o son `auto`, la celda combinada se queda sin ancho preferido y manda
  la rejilla.
- **Guardas**: las combinaciones rectangulares y las que romperían una combinación existente se
  rechazan con `InvalidOperationException` sin escribir nada; un span mal formado (sin `w:val`, no
  entero, `< 1`) lanza `InvalidOfficeFileException`.

**Fuera del 2.5c** (y por tanto no afirmado): las combinaciones **rectangulares** (se rechazan), el
`w:hMerge` heredado (se preserva) y los **márgenes de celda** (`w:tblCellMar`/`w:tcMar`, siguen
preservándose sin modelar).

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release` | **518 superadas, 0 con error, 0 omitidas** (518 total). |
| `dotnet test OfficeEngine.sln -c Release --filter "FullyQualifiedName~WordTableMerge"` | **46 superadas** (las nuevas del incremento). |

Salida textual del resumen de `dotnet test` (las duraciones varían entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   518, Omitido:     0, Total:   518
```

El informe TRX de Release está en
`artifacts/testresults/phase2/inc25c/coverage/phase2_inc25c_release.trx` y la cobertura en
`artifacts/testresults/phase2/inc25c/coverage/**/coverage.cobertura.xml`.

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Word/Table.cs` (acumulado 2.5a+2.5b+2.5c) | 163 / 163 | 100 % | 94,8 % (91/96) |
| `Word/TableCell.cs` (acumulado 2.5a+2.5b+2.5c) | 143 / 143 | 100 % | 97,7 % (84/86) |
| `Word/TableRow.cs` (acumulado 2.5a+2.5c) | 70 / 71 | 98,6 % | 94,7 % (36/38) |
| `Word/TableVerticalMerge.cs` | — | — | Enumeración: sin líneas ejecutables |
| **Total de la solución** | 2264 / 2353 | 96,2 % | 90,9 % (1321/1453) |

La única línea sin cubrir de los archivos tocados es el `_ => null` del `switch` que suma los anchos
al combinar en horizontal (`TableRow.MergedWidth`, línea 189), una salvaguarda inalcanzable: el bucle
devuelve antes si algún ancho es `auto` o falta, así que el tipo acumulado nunca es nulo al llegar ahí.
Las ramas parciales que quedan en `Table.cs` y `TableCell.cs` son las de 2.5b (`Properties?.…`), no de
este incremento: el código nuevo queda con cobertura de líneas y ramas completa.

### Pruebas nuevas del incremento (46 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordTableMergeTests.cs` (41) | Lectura: una tabla nueva no combina nada; `w:gridSpan` con valor 2/3/1; sin `w:val`, con valor 0, negativo o no numérico (error tipado); `w:vMerge` con `restart`, con `w:val` explícito de `continue` y sin `w:val` (se lee como `continue`); valor de combinación desconocido. Escritura: span en orden de esquema (tras `tcW`), asignar 1 elimina el elemento y poda `w:tcPr`, span fuera de rango, valor de combinación no definido, `w:vMerge` explícito y borrado al asignar `null`. Combinar en horizontal: la primera celda absorbe el rango y el contenido llega en orden, suma de spans ya combinados, suma de anchos del mismo tipo (puntos y porcentaje), sin ancho cuando no se puede sumar (falta o tipos mezclados), rechazo si alguna celda está en una combinación vertical, argumentos fuera de rango, conservación de una tabla anidada y del marcado no modelado (`w:noWrap`), división de la celda (contenido y celdas vacías nuevas), span almacenado de 1, celda de una sola columna, argumentos fuera de rango y rechazo con combinación vertical. Combinar en vertical: marcas `restart`/`continue`, contenido intacto, idempotencia, rechazo si la celda de arriba continúa una combinación, si los spans difieren y con argumentos fuera de rango; división vertical (limpieza de marcas y argumentos fuera de rango). Round-trip por stream de una tabla con combinación horizontal y vertical. |
| `Word/WordTableMergeInteropTests.cs` (5) | Interoperabilidad con el SDK: paquete válido con una fila combinada y una columna combinada en vertical, con lectura de `w:gridSpan`/`w:vMerge` y comprobación de que la suma de spans de cada fila sigue cuadrando con la rejilla; los dos estados de `TableVerticalMerge` producen documentos válidos y legibles; una fila combinada tiene una celda menos por columna absorbida; editar una tabla ajena conserva `w:hMerge` y `w:noWrap`; tras dividir, una celda por columna de la rejilla sin `w:gridSpan` y con el contenido repartido. |
| `Support/OoxmlValidation.cs` (+1 ayudante) | Fixture de una tabla de dos columnas y tres filas con la primera columna combinada en vertical (`restart`, sin `w:val` y `continue` explícito), construido con el SDK. |

Las 472 pruebas de los incrementos anteriores **siguen pasando** sin cambios: no hay regresiones.

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con celdas combinadas en las dos
   direcciones, de las tablas divididas y de los documentos ajenos editados.
2. **Lectura con el SDK** de las marcas escritas (`w:gridSpan`, `w:vMerge`), del número de celdas de
   cada fila y de la suma de spans frente a la rejilla, además del contenido repartido al dividir.
3. **Fixtures creados con el SDK** (span sin `w:val`, span de 1, valor de combinación desconocido,
   tabla con `w:hMerge` y `w:noWrap`) leídos por OfficeEngine, que debe leer el valor, fallar con
   error tipado o preservar lo que no modela.

### Hallazgos durante las pruebas y acciones tomadas

1. **Dos elementos vecinos, con opcionalidad opuesta.** El esquema **exige** `w:val` en `w:gridSpan`
   (sin él, el validador da error) pero lo hace **opcional** en `w:vMerge`, donde ausente significa
   `continue`. Por eso un span sin `w:val` es `InvalidOfficeFileException` y una combinación sin
   `w:val` se lee como `Continue`; al escribir, el valor de `w:vMerge` se declara siempre.
2. **La rejilla sigue siendo la fuente de verdad.** Una fila con menos celdas que columnas declara la
   rejilla es válida y el validador **no** comprueba que los spans sumen: la prueba de interoperabilidad
   comprueba esa suma por su cuenta, en vez de confiar en que el SDK la detecte.
3. **Combinar en horizontal no puede borrar celdas sin más.** Eliminar las celdas cubiertas destruiría
   su contenido en silencio, y la primera víctima sería lo que esta versión no modela: una tabla
   anidada. El combinado mueve todos los hijos (menos `w:tcPr`) y hay prueba de que la tabla anidada
   sobrevive y el paquete sigue validando.
4. **El combinado vertical no mueve el contenido.** La sonda confirmó que el esquema no prohíbe que
   una celda de continuación tenga párrafos, así que Word puede reconciliarlo; mover el texto sería
   decidir por el llamante sobre contenido ajeno. Queda documentado en el XML doc del método.
5. **El ancho de la celda combinada no se deja «a medias».** Dejar el ancho de la primera celda sería
   una pista incorrecta (media anchura para una celda que ocupa dos columnas). Se suma cuando se puede
   y se omite cuando no, para que mande la rejilla.
6. **Una prueba destapó una incoherencia real entre contrato e implementación.** `UnmergeCellsVertically`
   se implementó aceptando un tramo de **una** fila, mientras su documentación (y el diseño: es la
   inversa exacta de `MergeCellsVertically`) pedía dos; `UnmergeCellsVertically_OutOfRangeArguments_Throw`
   falló y se corrigió el contrato al más estricto, dejando el caso de una sola celda en
   `TableCell.VerticalMerge = null`, que ya existía.
7. **Sumar el ancho puede fallar al leer.** Combinar exige leer el ancho de cada celda cubierta: si
   alguna guarda un ancho que esta versión no modela (la unidad `nil`), la operación falla con
   `UnsupportedOfficeFeatureException` en vez de inventarse el ancho. Queda documentado en el XML doc
   de `MergeCells`.

### Qué **no** se ha implementado (y por tanto no se afirma)

- **Combinaciones rectangulares** (una fila combinada en horizontal que a su vez forma parte de una
  combinación vertical): se **rechazan** con `InvalidOperationException` y no se escribe nada, en vez
  de producir una tabla que Word no pueda reconciliar.
- **`w:hMerge`** (la marca heredada de combinación horizontal): se preserva, no se modela.
- **Márgenes de celda** (`w:tblCellMar`/`w:tcMar`): se preservan, no se modelan.
- **Combinación vertical con contenido**: el método no mueve ni borra el texto de las celdas que se
  combinan; si el llamante quiere que el área combinada muestre un único texto, lo mueve él.
- **Dividir una celda que pertenece a una combinación vertical**: se rechaza (rompería la
  combinación); **añadir o quitar filas o columnas** sigue fuera de alcance.
- La decisión de diseño (alcance, `ColumnSpan` como `int`, tri-state de `VerticalMerge`, contenido que
  viaja, guardas, ancho sumado) está en `docs/adr/0015-cell-merge.md`.

---

## Incremento 2.6 — Secciones, tamaño/orientación y márgenes

Fecha: 2026-10-19.

Modela la **sección** (`w:sectPr`) y su **configuración de página**. Hasta este incremento el
marcado se preservaba pero no se leía ni se escribía.

### Alcance entregado

- **Secciones en orden de documento** (`WordDocument.Sections`): el `w:sectPr` de cada párrafo que
  termina una sección y el `w:sectPr` final del cuerpo.
- **Añadir una sección** (`WordDocument.AddSection(startType)`): mueve las propiedades de la última
  sección a un párrafo de salto y deja una copia como `w:sectPr` final para la nueva, con el tipo de inicio
  indicado. No se pierde contenido.
- **Tamaño y orientación** (`Section.PageSize`, `PageSize`): `w:w`, `w:h` y `w:orient`, en puntos y
  con `Rotate()` para intercambiar las dimensiones y alternar la orientación.
- **Márgenes** (`Section.Margins`, `PageMargins`): `w:top`, `w:right`, `w:bottom` y `w:left`, en
  puntos. `w:header`, `w:footer`, `w:gutter` y el resto de atributos de `w:pgMar` se conservan.
- **Tipo de inicio** (`Section.StartType`): `nextPage`, `continuous`, `evenPage`, `oddPage` y
  `nextColumn` (`w:type`).
- **Sección por defecto** en `WordDocument.Create`: US Letter vertical, márgenes de una pulgada y
  distancias de encabezado/pie de 720 twips.
- Inserción en **orden de esquema** de `CT_SectPr` y conservación del `w:sectPr` final como último
  hijo del cuerpo (`AddParagraph` y `AddTable` insertan antes de él).

**Fuera del 2.6** (y por tanto no afirmado): encabezados/pies por sección (2.7), columnas, bordes de
página, numeración de líneas y de páginas, cuadrícula de documento y el resto de hijos de
`w:sectPr` (se preservan); borrar, mover o fusionar secciones.

### Decisiones de diseño

1. **Grupos atómicos.** `PageSize` y `PageMargins` son structs inmutables que se asignan como un
   todo. El esquema exige ambos atributos dentro de `w:pgSz` y los cuatro márgenes dentro de
   `w:pgMar`; escribir «solo el ancho» o «solo el margen izquierdo» dejaría markup inválido. `null`
   quita el elemento completo.
2. **La orientación no intercambia dimensiones.** Se escribe en `w:orient`; para el caso habitual de
   pasar de vertical a horizontal existe `PageSize.Rotate()`.
3. **US Letter por defecto.** Un documento nuevo es determinista en cualquier equipo; `PageSize.A4`
   está disponible para cambiarlo. Un documento abierto sin `w:sectPr` informa cero secciones y
   `AddSection` crea la sección por defecto sin añadir un salto.
4. **El salto de sección es un párrafo vacío**, que es como lo guarda Word; por eso aparece en
   `Paragraphs` y en `Text` como una línea vacía.
5. **Lectura estricta.** `w:pgSz` sin `w:w`/`w:h`, `w:pgMar` sin uno de los cuatro márgenes,
   `w:type` sin `w:val` o una medida/orientación no numérica lanzan `InvalidOfficeFileException`.
   Un valor definido por OOXML pero no modelado lanza `UnsupportedOfficeFeatureException` y el
   marcado se conserva. La escritura valida antes de tocar el árbol.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release` | **544 superadas, 0 con error, 0 omitidas** (544 total). |

Salida textual del resumen de `dotnet test` (la duración varía entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   544, Omitido:     0, Total:   544
```

Las **518 pruebas anteriores siguen pasando**. Solo se actualizaron dos aserciones que contaban los
hijos directos del cuerpo (`WordTableTests.AddTable_AppendsTheTableAtTheEndOfTheBody` y
`WordTableInteropTests.ProducedDocument_InterleavesParagraphsAndTables`) para incluir el `w:sectPr`
final que ahora escribe `WordDocument.Create`; el contenido y el orden de los bloques no cambian.

### Pruebas nuevas del incremento (26 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordSectionTests.cs` (17) | Sección por defecto (Letter, vertical, una pulgada); lectura de dos secciones en orden; tamaño: asignar, redondear a twips, orientación, round-trip por stream y por archivo y borrado; `Rotate()`; validación de dimensiones y enums; márgenes: asignar, leer, borrar, escribir en twips, admitir top/bottom negativos y rechazar horizontales negativos; `w:pgMar` sin margen requerido; los cinco tipos de inicio con round-trip y borrado; valor de enum no definido sin tocar el documento; `AddSection` con tamaño/márgenes independientes y round-trip; `AddSection` sobre un documento sin `w:sectPr`; argumentos fuera de rango; objeto dispuesto. |
| `Word/WordSectionInteropTests.cs` (9) | Paquete válido con `type`/`pgSz`/`pgMar` en orden y valores en twips; dos secciones con el `w:sectPr` final en su sitio; conservación de `w:cols`, `w:titlePg`, `w:docGrid`, `w:header`, `w:footer` y `w:gutter` al editar la configuración; editar una sección no toca la otra; `w:pgSz` con medida no numérica; orientación no válida; `w:pgMar` con medida no numérica; `w:type` sin `w:val` y con valor no válido. |
| `Support/OoxmlValidation.cs` (+2 ayudantes) | Lectura de las secciones en orden de documento y del `w:sectPr` final con el SDK, para que las aserciones no pasen por el código bajo prueba. |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con una y dos secciones, de los
   documentos ajenos editados y de los documentos con la configuración borrada.
2. **Lectura con el SDK** de `w:pgSz` (`w:w`, `w:h`, `w:orient`), `w:pgMar` (los cuatro márgenes y
   los atributos preservados), `w:type` y del orden de los hijos de `w:sectPr`.
3. **Fixtures creados con el SDK** (dos secciones, markup no modelado y valores mal formados) leídos
   por OfficeEngine, que debe leerlos, preservarlos o fallar con error tipado.

### Hallazgos durante las pruebas y acciones tomadas

1. **El `w:sectPr` final debe ser el último hijo del cuerpo.** Al escribir una sección por defecto,
   `AddParagraph` y `AddTable` habrían dejado el `w:sectPr` en medio y el paquete habría pasado a ser
   inválido. Se añadió un helper que inserta los bloques antes de él.
2. **Una prueba destapó un duplicado real.** El setter de `StartType` insertaba un `w:type` nuevo sin
   quitar el anterior; el getter devolvía el primero y el valor parecía no cambiar. Se corrigió
   eliminando el elemento previo antes de insertar, y la prueba de los cinco valores lo cubre.
3. **El SDK no expone `PageSize`/`PageMargin` como propiedades tipadas de `SectionProperties`.** La
   implementación y las pruebas usan `GetFirstChild<T>()`; la prueba de interoperabilidad verifica el
   orden y los valores con el propio SDK.
4. **`w:pgMar` exige los cuatro márgenes.** Por eso los márgenes se modelan como un grupo atómico y
   no como cuatro propiedades independientes que podrían dejar el elemento incompleto.
5. **La forma del cuerpo cambió para los documentos nuevos.** `WordDocument.Create` ahora produce
   `w:body` con un `w:sectPr`; se actualizaron las dos aserciones afectadas y se documentó en el ADR.

La decisión completa (alcance, valores por defecto, grupos atómicos, orden de esquema y
limitaciones) está en `docs/adr/0016-sections-and-page-setup.md`.

---

## Incremento 2.7 — Encabezados y pies por sección

Fecha: 2026-10-19.

Modela los **encabezados y pies** como partes del paquete referenciadas por sección. Hasta este
incremento el `w:sectPr` conservaba las referencias, pero no había forma de leer ni editar la parte que
nombraban.

### Alcance entregado

- **Lectura por sección y tipo** (`Section.Header(type)` / `Section.Footer(type)`,
  `HeaderFooterType`): `default`, `first` y `even`. `null` significa que la sección no declara ese tipo
  y hereda el de la sección anterior («vincular con la anterior»).
- **Creación de la referencia y la parte** (`Section.AddHeader` / `AddFooter`): una parte nueva con un
  párrafo vacío, o la existente si ya estaba declarada (idempotente).
- **Eliminación de la referencia** (`Section.RemoveHeader` / `RemoveFooter`) sin borrar la parte: otra
  sección puede seguir compartiéndola.
- **Contenido editable** (`HeaderFooter`): `Paragraphs`, `Tables`, `Text`, `AddParagraph` y `AddTable`,
  reutilizando los tipos `Paragraph` y `Table` del cuerpo y la misma validación de texto.
- **Vínculo sección-parte compartido**: varias secciones pueden apuntar a la misma parte; editar desde
  una se cambia lo que ven las demás. `WordDocument.AddSection` clona las referencias, así que la nueva
  sección comparte el encabezado y el pie de la anterior.
- **Primera página distinta** (`Section.DifferentFirstPage`, `w:titlePg`).

**Fuera del 2.7** (y por tanto no afirmado): `w:evenAndOddHeaders` (sin él Word no muestra el tipo
`EvenPage`; se preserva), el contenido no modelado de una parte (entonces imágenes y controles de
contenido; las imágenes inline se modelaron en 2.8) y los encabezados/pies por documento fuera de las
secciones.

### Decisiones de diseño

1. **Quitar la referencia no borra la parte.** Borrarla destruiría contenido compartido y el formato no
   exige que toda parte esté referenciada. La parte se conserva; los vínculos entre sección y parte no
   se rompen.
2. **Crear no reemplaza.** `AddHeader`/`AddFooter` solo crean cuando la sección no declara el tipo; si
   lo declara, devuelven la parte existente. Llamar dos veces no duplica referencias.
3. **Una parte nueva nace con un párrafo vacío**, que es lo que exige el esquema de `w:hdr`/`w:ftr` y
   lo que Word escribe. El llamante puede escribir en `Paragraphs[0]` o añadir otra línea con
   `AddParagraph` (por eso `Text` es `""` al crear y `"\n..."` tras añadir).
4. **Lectura estricta.** Una referencia sin `w:type`, con un valor no legible, sin `r:id`, con una
   relación ausente o que apunta a una parte del tipo equivocado lanza `InvalidOfficeFileException`;
   `w:titlePg` no booleano también. Un tipo no definido en un setter lanza
   `ArgumentOutOfRangeException` antes de tocar el árbol.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release` | **569 superadas, 0 con error, 0 omitidas** (569 total). |

Salida textual del resumen de `dotnet test` (la duración varía entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   569, Omitido:     0, Total:   569
```

Las **544 pruebas anteriores siguen pasando**. La refactorización de `WordDocument.Text` para compartir
el lector de bloques con `HeaderFooter.Text` no cambió el texto que informa.

### Pruebas nuevas del incremento (25 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordHeaderFooterTests.cs` (17) | Sin encabezados ni pies al crear; alta del encabezado con párrafo vacío y referencia; edición por párrafos; tabla en el pie; alta idempotente; los tres tipos con round-trip por stream, round-trip por archivo y borrado; `DifferentFirstPage` con round-trip y borrado; quitar la referencia conservando la parte; sección nueva que comparte el encabezado anterior; encabezados independientes en dos secciones; tipo no definido; referencia sin `w:type`, a parte ausente o con tipo mal formado; `w:titlePg` mal formado; validación de texto y dimensiones al añadir contenido. |
| `Word/WordHeaderFooterInteropTests.cs` (8) | Paquete válido con referencias y partes enlazadas; los tres tipos con el orden de referencias (`headerReference` antes de `footerReference`, y `pgSz`/`pgMar` detrás); dos secciones que comparten una parte y ven el mismo cambio; contenido no modelado (`w:sdt`) preservado al editar el texto; parte huérfana preservada; quitar la referencia de una sección y conservar la parte para la otra; referencia de encabezado a una parte de pie que falla con error tipado; paquete válido tras quitar todas las referencias. |
| `Support/OoxmlValidation.cs` (+10 ayudantes) | Fixtures con encabezados/pies por tipo, dos secciones que comparten una parte, parte huérfana, referencia a una parte del tipo equivocado y contenido no modelado; lectura independiente de recuento, texto y control de contenido. |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con encabezados y pies, de los
   documentos ajenos editados y de los documentos tras quitar referencias.
2. **Lectura con el SDK** de las referencias (`w:type` y `r:id`), de las partes (`HeaderParts` /
   `FooterParts`), de su texto y de las estructuras no modeladas.
3. **Fixtures creados con el SDK** (dos secciones que comparten una parte, parte huérfana, contenido no
   modelado y referencias mal formadas) leídos por OfficeEngine, que debe leerlos, preservarlos o
   fallar con error tipado.

### Hallazgos durante las pruebas y acciones tomadas

1. **Una parte es un objeto del paquete, no del documento.** Las primeras pruebas leían el texto de una
   parte después de cerrar el paquete y el SDK lanzaba `ObjectDisposedException`; los ayudantes de
   verificación leen ahora dentro del `using` del paquete.
2. **`AddHeader` crea un párrafo vacío.** El esquema exige al menos un bloque; las pruebas ajustaron sus
   expectativas (`Paragraphs[0].Text` para el primer párrafo, `Text` con una línea inicial vacía al
   añadir con `AddParagraph`).
3. **El orden de las referencias importa.** `w:headerReference` debe preceder a `w:footerReference`;
   la inserción por orden de esquema lo garantiza y la prueba de interoperabilidad lo verifica.
4. **Una vista `Section` anterior a `AddSection` apuntaba a la sección equivocada.** La prueba de dos
   secciones con encabezados independientes destapó que `AddSection` reutilizaba el `w:sectPr` del
   cuerpo para la sección nueva, mientras la sección que terminaba quedaba en una copia. Una vista
   `Section` obtenida antes de `AddSection` pasaba así a describir la sección nueva. Se corrigió
   moviendo el elemento original al párrafo de salto y dejando una copia para la nueva sección; la
   decisión quedó anotada en el ADR 0016 y la prueba nueva la cubre.

La decisión completa está en `docs/adr/0017-headers-and-footers.md`.

---

## Incremento 2.8 — Imágenes y relaciones

Fecha: 2026-10-19.

Modela las **imágenes inline incrustadas** y deja constancia de que las relaciones que la versión no
modela no se rompen. Hasta este incremento una imagen se preservaba porque el SDK conserva el
paquete, pero no se leía ni se creaba.

### Alcance entregado

- **Añadir una imagen inline** (`Paragraph.AddImage`):
  - desde archivo, infiriendo el formato de la extensión;
  - desde stream, con el formato declarado (`ImageFormat`: `Png`, `Jpeg`, `Gif`, `Bmp`, `Tiff`,
    `Emf` y `Wmf`).
  El tamaño se indica en **puntos** y se convierte a EMU (1 pt = 12 700 EMU); la imagen se añade al
  final del párrafo en un **run nuevo**, sin tocar su texto ni sus runs.
- **Leer las imágenes inline** (`Paragraph.Images`, `Image`): tamaño en puntos, tipo de contenido,
  identificador de relación y bytes (`ReadAllBytes`).
- **Cualquier parte que contenga párrafos**: cuerpo, celdas de tabla y encabezados/pies. La parte
  dueña se descubre desde el elemento, así que la relación se crea en `MainDocumentPart`,
  `HeaderPart` o `FooterPart` según corresponda.
- **Identificadores de dibujo únicos** (`wp:docPr/@id`): se toma el mayor de cuerpo, encabezados y
  pies y se usa el siguiente.
- **Lectura estricta**: una imagen inline mal formada (sin `wp:extent`, con extensión no positiva,
  sin `a:blip`, sin `r:embed`, con relación ausente o que no es `ImagePart`) lanza
  `InvalidOfficeFileException`.
- **Relaciones ajenas**: el SDK conserva las partes y relaciones existentes al guardar; hay pruebas
  de que editar el cuerpo o añadir una imagen no rompe un hipervínculo externo ni la relación de una
  imagen ajena.

**Fuera del 2.8** (y por tanto no afirmado): imágenes flotantes (`wp:anchor`), imágenes vinculadas
(`r:link`), gráficos y otros DrawingML, SVG/JPEG 2000/PCX/iconos al añadir, texto alternativo,
tamaño automático a partir de la imagen, reemplazo/eliminación/movimiento de imágenes, deduplicación
de partes y validación de que los bytes correspondan al formato declarado.

### Decisiones de diseño

1. **El tamaño se indica, no se deduce.** Decodificar cada formato para leer sus dimensiones añadiría
   dependencias o parsers propios; el formato tampoco lo exige. El llamante da puntos y la API los
   convierte a EMU validando antes de crear la parte.
2. **Cada llamada crea su propia parte.** No se deduplican bytes; es válido y predecible. Queda como
   optimización futura.
3. **La parte dueña se descubre desde el elemento.** Subir por los ancestros hasta el
   `OpenXmlPartRootElement` evita enhebrar la parte por `Paragraph`/`Table`/`TableCell`/`HeaderFooter`
   y hace que la misma API funcione en el cuerpo y en las historias de encabezado/pie.
4. **Al leer, solo se modela lo que se puede editar con seguridad.** Las imágenes flotantes,
   vinculadas y los gráficos se preservan y no se devuelven; una imagen inline incrustada mal formada
   falla con error tipado en vez de inventar dimensiones.
5. **No se reescribe el paquete.** Las relaciones existentes se conservan; la API solo añade la parte
   y la relación de la imagen nueva.

La decisión completa está en `docs/adr/0018-images-and-relationships.md`.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release` | **594 superadas, 0 con error, 0 omitidas** (594 total). |
| `dotnet test OfficeEngine.sln -c Release --filter "FullyQualifiedName~WordImage"` | **25 superadas** (las nuevas del incremento). |

Salida textual del resumen de `dotnet test` (la duración varía entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   594, Omitido:     0, Total:   594
```

Las **569 pruebas anteriores siguen pasando** sin cambios.

### Pruebas nuevas del incremento (25 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordImageTests.cs` (21) | Un documento nuevo no tiene imágenes; añadir desde archivo con round-trip por archivo (tipo de contenido, tamaño, bytes y conservación del texto); añadir desde stream y conversión de puntos a EMU; los siete formatos de `ImageFormat` producen su tipo de contenido; la imagen se añade en un run nuevo sin cambiar el texto; imagen en celda de tabla con round-trip; imagen en encabezado con round-trip; dos imágenes con identificadores de dibujo distintos; validación de ruta nula/vacía y extensión no soportada, stream nulo o no legible, formato no definido y dimensiones cero/negativas/`NaN`/infinitas; una dimensión inválida no deja parte de imagen; lectura de una imagen ajena construida con el SDK; error tipado cuando la relación no existe, cuando falta `wp:extent` y cuando la relación apunta a una parte que no es una imagen. |
| `Word/WordImageInteropTests.cs` (4) | Paquete válido con una imagen y relación resoluble; una imagen ajena conserva bytes, relación e identificador al editar el cuerpo; un hipervínculo externo y su `r:id` sobreviven a añadir una imagen; imágenes en cuerpo, tabla, encabezado y pie dan un paquete válido con partes por historia e identificadores únicos. |
| `Support/OoxmlValidation.cs` (+6 fixtures y 9 ayudantes) | Fixtures de imagen inline con PNG, imagen sin `wp:extent`, relación a una parte que no es imagen, relación a una parte ausente e hipervínculo externo; recuento y lectura de partes de imagen, identificadores de dibujo, extensiones inline y relaciones de hipervínculo con el SDK. |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con una y varias imágenes, de los
   documentos ajenos editados y de los documentos con imágenes en distintas historias.
2. **Lectura con el SDK** de `ImageParts`, tipo de contenido, bytes, `wp:docPr/@id`, `wp:inline` y de
   la relación de hipervínculo externa, además de resolver `a:blip/@r:embed` contra la parte.
3. **Fixtures creados con el SDK** (imagen inline, imagen sin `wp:extent`, relación a una parte que
   no es imagen, relación ausente e hipervínculo externo) leídos por OfficeEngine, que debe leerlos,
   preservarlos o fallar con error tipado.

### Hallazgos durante las pruebas y acciones tomadas

1. **`ImagePartType` es una clase estática con `PartTypeInfo`, y `AddImagePart` solo existe en las
   partes que declaran la relación.** No se puede llamar sobre un `OpenXmlPart` cualquiera; el tipo
   de la parte se comprueba explícitamente para `MainDocumentPart`, `HeaderPart` y `FooterPart`.
2. **La parte dueña se descubre desde el elemento.** Al principio el diseño preveía enhebrar la parte
   por los envoltorios; subir por los ancestros hasta `OpenXmlPartRootElement` fue más simple y
   además hizo que las imágenes en celdas y encabezados funcionaran sin cambios.
3. **Reutilizar un stream no añade una segunda imagen.** `FeedData` copia desde la posición actual
   hasta el final; las pruebas crean un stream nuevo (o reinician la posición) por imagen. Queda como
   responsabilidad del llamante, igual que en el resto de la API.
4. **Un id de dibujo duplicado es un riesgo silencioso.** Se toma el mayor de cuerpo, encabezados y
   pies para que dos imágenes añadidas en historias distintas no compartan `wp:docPr/@id`.
5. **Conservar lo que no se modela se prueba con una relación ajena, no con una promesa.** El fixture
   de hipervínculo externo comprueba que añadir una imagen no toca una relación que la API no
   modela, que es el criterio del incremento.

### Qué **no** se ha implementado (y por tanto no se afirma)

- **Imágenes flotantes y vinculadas**: se preservan, no se leen ni se crean.
- **Gráficos, SmartArt y otros DrawingML**: se preservan, no se modelan.
- **Formatos SVG, JPEG 2000, PCX e iconos al añadir**: `ImageFormat` no los ofrece; si ya existen en
  el documento, se preservan.
- **Texto alternativo, tamaño automático, deduplicación, reemplazo, eliminación y movimiento**:
  fuera de alcance.
- **Validación de que los bytes correspondan al formato declarado**: la API copia lo que recibe y
  fija el tipo de contenido; no decodifica la imagen.
- La decisión completa está en `docs/adr/0018-images-and-relationships.md`.

---

## Incremento 2.9 — Hipervínculos externos

Fecha: 2026-10-19.

Modela los **hipervínculos externos** y deja constancia de que los enlaces que la versión no modela
se preservan. Hasta este incremento un hipervínculo se conservaba porque el SDK conserva el paquete,
pero no se leía ni se creaba. Un enlace externo es un `w:hyperlink` con un `r:id` que apunta a una
relación de hipervínculo de la parte que contiene el párrafo.

### Alcance entregado

- **Crear un hipervínculo externo** (`Paragraph.AddHyperlink(text, Uri)`): añade un `w:hyperlink`
  con una relación nueva al final del párrafo, sin tocar su texto ni sus runs directos.
- **Leer los hipervínculos externos** (`Paragraph.Hyperlinks`, `Hyperlink`): texto visible
  (`Text`), runs directos (`Runs`) y destino (`Destination`).
- **Editar el texto y el destino**: `Hyperlink.Text` reemplaza el contenido por un run nuevo y
  conserva el destino; `Hyperlink.Destination` crea una relación nueva y actualiza el `r:id`.
- **Cualquier parte que contenga párrafos**: cuerpo, celdas de tabla y encabezados/pies. La parte
  dueña se descubre desde el elemento, así que la relación se crea en `MainDocumentPart`,
  `HeaderPart` o `FooterPart` según corresponda.
- **Relaciones compartidas**: si varios enlaces comparten una relación, editar uno no cambia los
  demás; la relación anterior se elimina solo cuando queda sin uso en la parte.
- **Lectura estricta**: un enlace con `r:id` cuya relación no existe, no es de hipervínculo o no es
  externa lanza `InvalidOfficeFileException`. Un enlace interno (`w:anchor`) se preserva y no se
  devuelve.
- **Relaciones ajenas**: se conservan al guardar; las pruebas comprueban que los enlaces internos y
  las relaciones compartidas sobreviven a las ediciones.

**Fuera del 2.9** (y por tanto no afirmado): enlaces internos a marcadores (`w:anchor`), creación o
edición de `w:tooltip`, `w:history`, `w:tgtFrame` o `w:docLocation`, estilo o formato automático del
texto del enlace, eliminación o movimiento de enlaces, y validación del esquema de la `Uri` (la API
almacena el destino y no lo abre ni lo ejecuta).

### Decisiones de diseño

1. **Solo enlaces externos.** Un enlace interno no tiene relación: su destino es un marcador del
   documento. Modelarlo exige definir cómo se leen, crean y editan los marcadores, que aún no
   forman parte de la API. Se acota el incremento y se preserva el `w:anchor`.
2. **El destino vive en la relación.** Es como lo almacena Word y permite destinos relativos o
   absolutos sin inventar un formato de texto propio.
3. **Las relaciones no se comparten por accidente.** Editar un enlace crea una relación nueva para él
   y deja intacta la relación compartida por otros; la anterior se borra solo si ya no la usa nadie.
4. **El texto se puede editar con o sin formato.** `Hyperlink.Text` es la vía simple y documenta que
   reemplaza el contenido; `Hyperlink.Runs` permite conservar el formato de un run existente.
5. **No se aplica el estilo `Hyperlink`.** Un documento nuevo de OfficeEngine no define ese estilo y
   el formato visual no forma parte del criterio del incremento; el llamante puede aplicar formato a
   `Hyperlink.Runs`.

La decisión completa está en `docs/adr/0019-external-hyperlinks.md`.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **618 superadas, 0 con error, 0 omitidas** (618 total). |
| `dotnet test OfficeEngine.sln -c Release --no-build --filter "FullyQualifiedName~WordHyperlink"` | **24 superadas** (las nuevas del incremento). |

Salida textual del resumen de las pruebas (la duración varía entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   618, Omitido:     0, Total:   618
```

Las **594 pruebas anteriores siguen pasando** sin cambios.

### Pruebas nuevas del incremento (24 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordHyperlinkTests.cs` (18) | Un documento nuevo no tiene enlaces; creación y round-trip por archivo (texto, destino, `Paragraph.Text`, runs directos); round-trip por stream; enlace en celda de tabla; enlaces en encabezado y pie; dos enlaces con relaciones distintas aunque el destino coincida; edición de texto conservando el destino; `Text` descarta los runs anteriores pero conserva la relación; edición del texto de un run conservando su formato; edición del destino y eliminación de la relación anterior; relación compartida sin alterar el otro enlace; lectura de un enlace ajeno; enlace interno preservado y no devuelto; error tipado ante relación ausente, que no es de hipervínculo o que no es externa; validación de argumentos y de ediciones inválidas sin cambiar el enlace. |
| `Word/WordHyperlinkInteropTests.cs` (6) | Paquete válido con enlace y relación resoluble; lectura y edición de texto y destino con fixture ajeno; enlace interno preservado al editar el cuerpo; enlaces en cuerpo, tabla, encabezado y pie con relaciones en la parte correcta; relación compartida que sigue en uso tras editar un enlace; edición del destino de un enlace en encabezado y limpieza de la relación anterior. |
| `Support/OoxmlValidation.cs` (+5 fixtures y 5 ayudantes) | Fixtures de dos enlaces que comparten relación, enlace interno, relación ausente, relación a una parte que no es hipervínculo y relación de hipervínculo no externa; lectura de textos, anclas y destinos de los enlaces, y de las relaciones externas de encabezados y pies con el SDK. |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con enlaces creados, editados y
   repartidos entre cuerpo, tabla, encabezado y pie.
2. **Lectura con el SDK** de los `w:hyperlink`, sus `r:id`, las `HyperlinkRelationship` de cada parte,
   su `IsExternal` y su `Uri`, además del texto visible.
3. **Fixtures creados con el SDK** (enlace externo, dos enlaces que comparten relación, enlace interno,
   relación ausente y relación a una parte que no es hipervínculo) leídos por OfficeEngine, que debe
   leerlos, preservarlos o fallar con error tipado.

### Hallazgos durante las pruebas y acciones tomadas

1. **Los identificadores de relación son locales a la parte.** Al editar un destino hay que buscar la
   relación en la parte que contiene el enlace, no en el documento principal; por eso la misma API
   funciona en encabezados y pies.
2. **Eliminar una relación compartida rompería otro enlace.** Antes de borrar la relación anterior se
   recorre la parte para comprobar que ningún `w:hyperlink` sigue usando ese `r:id`; la prueba con dos
   enlaces que comparten relación cubre el caso.
3. **Un enlace interno no debe confundirse con uno externo.** `Paragraph.Hyperlinks` filtra los que no
   tienen `r:id` y los preserva; la prueba comprueba que el `w:anchor` sigue en el paquete tras editar
   el cuerpo.
4. **`Paragraph.Runs` no devuelve los runs de un enlace.** Es coherente con el modelo: el enlace es un
   hijo del párrafo, no un run. `Hyperlink.Runs` es la vista para editarlos.

### Qué **no** se ha implementado (y por tanto no se afirma)

- **Enlaces internos a marcadores** (`w:anchor`): se preservan, no se leen ni se crean.
- **Atributos adicionales** (`w:tooltip`, `w:history`, `w:tgtFrame`, `w:docLocation`): se preservan,
  no se modelan.
- **Estilo o formato automático del enlace**: no se aplica; se puede usar `Hyperlink.Runs`.
- **Eliminar, mover o reordenar enlaces**: fuera de alcance, igual que en párrafos y runs.
- **Validación del destino**: no se comprueba el esquema de la `Uri` ni se abre.
- La decisión completa está en `docs/adr/0019-external-hyperlinks.md`.

---

## Incremento 2.10 - Propiedades básicas del documento

Fecha: 2026-10-19.

Modela las **propiedades básicas** que Word guarda en `docProps/core.xml` y deja constancia de que
las propiedades que la versión no modela se preservan. Hasta este incremento los metadatos se
conservaban porque el SDK conserva el paquete, pero no se leían ni se editaban.

### Alcance entregado

- **Leer y editar** título, asunto, autor, palabras clave, descripción, último editor, revisión,
  fecha de impresión, fecha de creación, fecha de modificación, categoría, estado, identificador,
  idioma y versión mediante `WordDocument.Properties` y `DocumentProperties`.
- **Vista viva**: los cambios se aplican al documento en memoria y se guardan con `Save`. Asignar
  `null` elimina la propiedad; leer una propiedad ausente devuelve `null`.
- **Parte convencional**: si el documento no tiene propiedades y se asigna la primera, se crea
  `docProps/core.xml` en la ruta que usa Word. Leer o borrar una propiedad inexistente no crea la
  parte.
- **Preservación**: `docProps/app.xml` (propiedades extendidas) y `docProps/custom.xml` (propiedades
  personalizadas) se conservan sin cambios al editar las básicas.
- **Validación de texto**: se permiten tabuladores y saltos de línea, pero un carácter que XML no
  puede almacenar se rechaza con `ArgumentException` antes de tocar la parte.
- **Lectura estricta**: una parte de core properties con una fecha mal formada lanza
  `InvalidOfficeFileException` en vez de una excepción del SDK.
- **Fechas**: se conserva el instante; el SDK las devuelve como `DateTime` local, así que el offset
  textual puede cambiar. La comparación correcta es por `ToUniversalTime`.

**Fuera del 2.10** (y por tanto no afirmado): propiedades extendidas (`docProps/app.xml`) y
personalizadas (`docProps/custom.xml`), estadísticas de la aplicación, empresa/gerente,
`ContentType`, edición de la parte de propiedades de una firma digital o de cualquier metadato que no
sea una propiedad básica. Esas partes se preservan pero no se leen ni se editan.

### Decisiones de diseño

1. **Solo propiedades básicas.** Son las que Word muestra en la ficha del documento y las que el
   criterio del incremento puede verificar con round-trip. Las extendidas y personalizadas tienen
   estructuras distintas y quedan fuera.
2. **`docProps/core.xml`, no una ruta OPC arbitraria.** `WordprocessingDocument.PackageProperties`
   crea una parte `package/services/metadata/core-properties/<guid>.psmdcp` cuando no existe. Para
   que el paquete producido siga la convención de Word, se crea un `CoreFilePropertiesPart` en
   `docProps/core.xml` y se inicializa su raíz vacía antes de escribir.
3. **El SDK sigue siendo el lector/escritor.** El SDK 3.5.1 expone las propiedades mediante
   `IPackageProperties`, marcada como experimental (`OOXML0001`). Se usa encapsulada, sin exponerla
   en la API pública y sin nombrar el tipo en el código de producción; la versión está fijada y las
   pruebas cubren el round-trip.
4. **Sin inventar valores.** Una propiedad ausente es `null`; una propiedad mal formada falla con
   error tipado. No se rellena una fecha o un texto por defecto.
5. **No se reescribe el XML a mano.** El SDK conoce los nombres de elemento, el orden y el tipo
   `dcterms:W3CDTF`, lo que reduce el riesgo de generar marcado inválido.

La decisión completa está en `docs/adr/0020-document-properties.md`.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **632 superadas, 0 con error, 0 omitidas** (632 total). |
| `dotnet test OfficeEngine.sln -c Release --no-build --filter "FullyQualifiedName~WordProperties"` | **14 superadas** (las nuevas del incremento). |

Salida textual del resumen de las pruebas (la duración varía entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   632, Omitido:     0, Total:   632
```

Las **618 pruebas anteriores siguen pasando** sin cambios.

### Pruebas nuevas del incremento (14 casos)

| Archivo | Foco |
| --- | --- |
| `Word/WordPropertiesTests.cs` (10) | Documento nuevo sin propiedades ni parte; borrar valores inexistentes sin crear la parte; round-trip de las quince propiedades por archivo; edición de valores existentes por stream; borrado de propiedades; preservación de `app.xml`/`custom.xml`; tabuladores y saltos de línea; rechazo de caracteres XML inválidos; fecha mal formada con `InvalidOfficeFileException`; uso tras `Dispose`. |
| `Word/WordPropertiesInteropTests.cs` (4) | Paquete válido y legible por el SDK; nombres, namespaces y `xsi:type` esperados; edición de un fixture ajeno sin perder partes no modeladas; creación de `docProps/core.xml` al editar un documento que no lo tenía. |
| `Support/OoxmlValidation.cs` (+2 fixtures y 4 ayudantes) | Fixture con core, extended y custom properties; fixture con una fecha mal formada; lectura del XML de core properties, de la empresa extendida y de una propiedad personalizada con el SDK. |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con propiedades creadas y
   editadas.
2. **Lectura con el SDK** de la parte de core properties, su ruta, sus valores, sus fechas y las
   partes extendida y personalizada.
3. **Análisis del XML** con `XDocument`: namespaces `dc`, `cp`, `dcterms` y `xsi`, nombres de
   elemento y `xsi:type="dcterms:W3CDTF"` en las fechas que lo requieren.
4. **Fixtures creados con el SDK** leídos y editados por OfficeEngine.

### Hallazgos durante las pruebas y acciones tomadas

1. **`PackageProperties` no usa `docProps/core.xml`.** En SDK 3.5.1 crea una parte `.psmdcp` en
   `package/services/metadata/core-properties/`. Se cambió a `CoreFilePropertiesPart` para producir
   la ruta convencional de Word y se inicializa una raíz vacía antes de escribir.
2. **`lastPrinted` no lleva `xsi:type`.** Es un `xsd:dateTime` de `cp`, a diferencia de `created` y
   `modified`; la prueba de XML lo refleja.
3. **Una fecha mal formada lanza `XmlException` desde el SDK.** Se envuelve como
   `InvalidOfficeFileException` para mantener el contrato de errores tipados.
4. **Escribir la raíz vacía es necesario.** Un `CoreFilePropertiesPart` recién añadido está vacío y
   el SDK no puede leerlo hasta que existe un elemento raíz; la creación se hace una sola vez y solo
   al asignar el primer valor.

### Qué **no** se ha implementado (y por tanto no se afirma)

- **Propiedades extendidas** (`docProps/app.xml`): se preservan, no se leen ni se editan.
- **Propiedades personalizadas** (`docProps/custom.xml`): se preservan, no se leen ni se editan.
- **`ContentType`**: no se expone porque describe la parte, no un metadato editable.
- **Fechas con offset textual estable**: se conserva el instante, no necesariamente el offset.
- La decisión completa está en `docs/adr/0020-document-properties.md`.


## Incremento 2.11 - Búsqueda y reemplazo

### Alcance

Buscar y reemplazar texto en el cuerpo de un documento de Word, con un alcance explícito y reglas
claras de preservación de formato. Incluye:

- `WordDocument.FindText(text, matchCase)` y `WordDocument.ReplaceText(oldText, newText, matchCase)`.
- `Paragraph.FindText` y `Paragraph.ReplaceText`, con el alcance limitado a un párrafo.
- `WordTextMatch`, con el párrafo, el desplazamiento, la longitud y el texto de la coincidencia.
- Búsqueda de texto repartido en varios runs y reemplazo conservando el formato del primero.

La decisión completa está en `docs/adr/0036-word-text-search-and-replace.md`.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **909 superadas, 0 con error, 0 omitidas** (909 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase2/inc211` | **909 superadas**; cobertura en `artifacts/testresults/phase2/inc211/**/coverage.cobertura.xml`. |

Las **885 pruebas anteriores siguen pasando**; las 16 nuevas están en `WordTextSearchTests`.

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Word/WordTextSearch.cs` | 71 / 74 | 96,0 % | 85,0 % |
| `Word/WordTextMatch.cs` | 9 / 11 | 81,8 % | 100 % |
| **Total de la solución** | 5557 / 5921 | 93,9 % | 82,3 % |

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` del paquete producido.
2. **Lectura con el SDK** del texto del documento producido.
3. **Round-trip** por guardado y reapertura: el texto reemplazado es el que se vuelve a leer.

### Hallazgos durante las pruebas y acciones tomadas

1. **El texto de un párrafo no es una cadena.** «borrador» puede estar en tres runs con formatos
   distintos. La búsqueda trabaja sobre la concatenación de los `w:t` y el reemplazo deja el texto
   nuevo en el primer run de la coincidencia, que es la regla de Word.
2. **El alcance tiene que ser explícito.** Se buscó primero en todo el documento y se vio que eso
   alcanzaba a los encabezados y a los pies; el alcance se fijó en el cuerpo (párrafos de primer
   nivel y párrafos de las tablas de primer nivel) y la búsqueda no cruza párrafos.
3. **Las coincidencias se aplican de la última a la primera.** Si se sustituyen en orden, los
   desplazamientos de las que quedan pendientes dejan de valer.

### Qué **no** se ha implementado (y por tanto no se afirma)

- Búsqueda en encabezados, pies, notas, comentarios y tablas anidadas.
- Búsqueda con expresiones regulares, búsqueda inversa y reemplazo con comodines.
- Búsqueda que cruce párrafos, y elección del formato que conserva el texto reemplazado (es siempre el
  del primer run).

## Incremento 2.12 - Campos y numeración de páginas

### Alcance

Escribir campos de Word, que es lo que permite la numeración de páginas, sin evaluarlos. Incluye:

- `Paragraph.AddField(instruction, cachedResult)`, que escribe la secuencia de runs del campo.
- `HeaderFooter.AddPageNumber(cachedResult)`, que escribe el campo `PAGE` en el pie.
- `WordField`, con la instrucción y el resultado almacenado.

La decisión completa está en `docs/adr/0037-word-fields-and-page-numbering.md`.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **909 superadas, 0 con error, 0 omitidas** (909 total). |

Las **901 pruebas anteriores siguen pasando**; las 8 nuevas están en `WordFieldTests`.

### Verificación independiente (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` del paquete con campos.
2. **Lectura con el SDK** de la instrucción (`w:instrText`) y de los caracteres de campo
   (`w:fldChar`) de la parte de pie (`OoxmlValidation.ReadDocxFooterFields`).
3. **Round-trip** por guardado y reapertura: el campo y su resultado siguen ahí.

### Hallazgos durante las pruebas y acciones tomadas

1. **Un campo son cinco runs, no un texto.** `begin`, la instrucción, `separate`, el resultado y
   `end`. El resultado es opcional, y sin él el campo se escribe con tres runs, que es la forma
   mínima válida.
2. **El resultado almacenado no se recalcula.** Es el que el llamante da; Word lo recalcula al abrir
   o al imprimir. Por eso la API no promete el número de página real.
3. **El número de página reutiliza el párrafo vacío de la parte.** Si no, el pie quedaría con una
   línea en blanco encima del número.
4. **La instrucción no es texto visible.** El texto del párrafo muestra el resultado del campo, no
   la instrucción, porque la instrucción es un `w:instrText` y no un `w:t`.

### Qué **no** se ha implementado (y por tanto no se afirma)

- Cálculo o actualización de los campos: la librería escribe campos, no los evalúa.
- `w:fldSimple`, campos anidados, campos con formato de resultado y la marca `dirty`.
- El resto de los campos comunes de Word (fecha, título, autor, número de páginas) solo se pueden
  escribir con `AddField`; no hay una API propia para cada uno.
