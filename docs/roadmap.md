# Hoja de ruta y criterios de aceptación

Cada incremento se considera terminado solo cuando cumple la *definición de terminado* de
`AGENT.md`: criterios satisfechos, compila, pruebas relevantes pasan, API documentada, matriz
actualizada, limitaciones documentadas y sin regresiones sin explicar.

## Definición de terminado (por incremento)

1. Contrato público claro y documentado con XML docs.
2. Pruebas unitarias y/o de integración apropiadas.
3. Casos de ida y vuelta relevantes pasan.
4. Verificación con una herramienta independiente cuando aplique.
5. Limitaciones y posibles pérdidas documentadas.
6. Fila correspondiente en `docs/compatibility-matrix.md` con estado real.
7. Build sin avisos (`TreatWarningsAsErrors`).

## Fase 0 — Descubrimiento y decisiones — **Completada**

Entregables: arquitectura (`docs/architecture/overview.md`), ADR 0001–0007, evaluación de
dependencias y licencias, matriz de compatibilidad inicial, hoja de ruta.

## Fase 1 — Base compilable — **Completada**

| # | Incremento | Criterio de aceptación | Estado |
| --- | --- | --- | --- |
| 1.1 | Solución, proyectos mínimos, CI, análisis estático | `dotnet build -c Release` sin avisos; CI definida; analizadores activos | Hecho |
| 1.2 | Errores tipados, límites, diagnósticos, streams | Pruebas de límites, diagnósticos y jerarquía de excepciones pasan | Hecho |
| 1.3 | Carga acotada e inspección del paquete | Rechazo de no-ZIP, exceso de tamaño/entradas/ratio con pruebas | Hecho |
| 1.4 | Crear/guardar/reabrir un DOCX mínimo | Round-trip de párrafos; validación con el SDK; lectura del XML crudo | Hecho |
| 1.5 | Crear/guardar/reabrir un XLSX mínimo | Round-trip de valores; validación con el SDK; lectura del XML crudo | Hecho |
| 1.6 | Conservación de contenido no modelado | Tabla y `mergeCells` sobreviven al round-trip y se reportan | Hecho |

Resultados reales: ver `docs/phase-1-report.md`.

## Fase 2 — Word avanzado inicial — **Completada**

Cada incremento exige round-trip, verificación con una herramienta independiente y
documentación de limitaciones.

| # | Incremento | Criterio de aceptación | Estado |
| --- | --- | --- | --- |
| 2.1 | Párrafos/runs y formato de carácter (fuente, tamaño, negrita, cursiva, subrayado, color) | Round-trip por archivo y stream; paquete válido con el SDK; `w:rPr` en orden de esquema; lo no modelado se preserva; formato leído desde un archivo ajeno; entradas ilegibles fallan con error tipado | Hecho |
| 2.2 | Formato de párrafo: alineación, sangría y espaciado | Round-trip por archivo y stream; paquete válido con el SDK; `w:pPr` en orden de esquema; valores no modelados leídos desde un archivo ajeno fallan con error tipado en vez de inventarse; lo no modelado se preserva | Hecho |
| 2.2b | Tabulaciones (`w:tabs`/`w:tab`) y saltos explícitos (`w:br`/`w:cr`) | Round-trip por archivo y stream; validación con el SDK; `w:tabs` en orden de esquema; posiciones, alineación y relleno de tabulación legibles; lo no modelado se preserva y lo ilegible falla con error tipado | Hecho |
| 2.3 | Estilos de párrafo y carácter | Round-trip; herencia de estilos respetada; limitaciones documentadas | Hecho |
| 2.4 | Listas numeradas y con viñetas | Round-trip; numeración legible y editable sin romper la lista | Hecho |
| 2.5a | Tablas: estructura y contenido (crear/leer/editar filas, celdas y su texto) | Round-trip por archivo y stream; paquete válido con el SDK; `w:tbl` en orden de esquema; tablas existentes no se destruyen al editar; lo no modelado se preserva | Hecho |
| 2.5b | Tablas: propiedades (ancho, alineación, disposición, bordes, sombreado y alineación vertical) | Round-trip por archivo y stream; paquete válido con el SDK; lo no modelado se preserva; un valor definido por OOXML que no se modela falla con error tipado en vez de inventarse | Hecho |
| 2.5c | Tablas: combinación de celdas (`w:gridSpan`/`w:vMerge`) | Round-trip; validación con el SDK; lo no modelado se preserva | Hecho |
| 2.6 | Secciones, tamaño/orientación y márgenes | Round-trip; validación con el SDK | Hecho |
| 2.7 | Encabezados y pies por sección | Round-trip; vínculos entre sección y parte conservados | Hecho |
| 2.8 | Imágenes y relaciones | Round-trip; relaciones no modeladas no se rompen | Hecho |
| 2.9 | Hipervínculos | Round-trip; texto y destino legibles y editables | Hecho |
| 2.10 | Propiedades del documento | Round-trip; propiedades no modeladas se preservan | Hecho |
| 2.11 | Búsqueda y reemplazo con alcance explícito y reglas de preservación de formato | Alcance definido y probado; coincidencias repartidas en varios runs | Hecho |
| 2.12 | Numeración de páginas y campos comunes | Round-trip; se documenta que los campos se generan, no se evalúan | Hecho |

Nota: el alcance previsto del 2.2 incluía tabulaciones y saltos. Al implementarlo se dividió en dos
incrementos (2.2 y 2.2b) **en lugar de declarar hecho algo a medias**; el informe de la Fase 2
documenta el alcance realmente entregado en cada uno.

Nota: el 2.11 fija el alcance de la búsqueda en el cuerpo (párrafos de primer nivel y párrafos de
las tablas de primer nivel) y el 2.12 escribe campos sin evaluarlos. Ver
`docs/adr/0036-word-text-search-and-replace.md` y `docs/adr/0037-word-fields-and-page-numbering.md`.

Nota: el 2.3 se acotó expresamente a **referencias de estilo** (`w:pStyle`/`w:rStyle`) y al
**catálogo de estilos de solo lectura** con resolución de `w:basedOn`. Crear, editar o eliminar
definiciones de estilo queda fuera de este incremento (ver `docs/adr/0011-style-references-and-catalogue.md`).

Nota: el 2.4 se acotó a **referencias de lista** (`w:numPr`), al **catálogo de listas de solo
lectura** y a la **creación de listas de un solo nivel** con `WordDocument.AddList`. Personalizar
niveles (texto, inicio, reinicios) queda fuera de este incremento (ver
`docs/adr/0012-lists-references-catalogue-and-creation.md`).

Nota: el 2.5 se dividió expresamente en **2.5a** (estructura y contenido), **2.5b** (propiedades) y
**2.5c** (combinación de celdas), igual que se hizo con el 2.2/2.2b, **en lugar de declarar hecho
algo a medias**. El 2.5a entrega tablas creadas y editables en su estructura y su texto, conservando
el formato que no modela; el 2.5b modela ese formato: ancho de tabla y de celda, alineación,
disposición (`w:tblLayout`), bordes, sombreado y alineación vertical, y deja fuera los **márgenes de
celda** (`w:tblCellMar`/`w:tcMar`), que se siguen preservando sin modelar (ver
`docs/adr/0013-tables-structure-and-content.md` y `docs/adr/0014-table-and-cell-properties.md`).
El 2.5c cierra el 2.5: las celdas combinadas se leen y se escriben (`w:gridSpan`/`w:vMerge`) y hay
operaciones de combinar y dividir horizontales y verticales que mueven o crean celdas sin descartar
contenido; las combinaciones rectangulares se rechazan con error tipado y `w:hMerge` heredado y los
márgenes de celda siguen preservándose sin modelar (ver `docs/adr/0015-cell-merge.md`).

Nota: el 2.6 modela **secciones** (`w:sectPr`), **tamaño/orientación** y **márgenes**, y permite
**añadir** una sección al final. Las columnas, los bordes de página y la numeración quedan fuera y se
preservan; borrar, mover o fusionar secciones tampoco se modela (ver
`docs/adr/0016-sections-and-page-setup.md`).

Nota: el 2.7 modela los **encabezados y pies por sección**: las referencias por tipo (`default`,
`first` y `even`), la creación y eliminación de la referencia conservando la parte, el contenido
(párrafos, tablas y texto) y el hecho de que varias secciones pueden **compartir** la misma parte.
`w:titlePg` se modela como `Section.DifferentFirstPage`; `w:evenAndOddHeaders` (necesario para que
Word muestre los pares) se preserva pero no se modela (ver `docs/adr/0017-headers-and-footers.md`).

Nota: el 2.8 modela las **imágenes inline incrustadas**: añadirlas desde archivo o stream
(`Paragraph.AddImage`, `ImageFormat`), leerlas (`Paragraph.Images`, `Image`) con tamaño en puntos,
tipo de contenido y bytes, en el cuerpo, celdas de tabla y encabezados/pies. Cada imagen crea su
propia parte y relación; las relaciones no modeladas (por ejemplo un hipervínculo externo) y las
imágenes flotantes, vinculadas o los gráficos se preservan sin modelarse (ver
`docs/adr/0018-images-and-relationships.md`).

Nota: el 2.9 modela los **hipervínculos externos**: crearlos (`Paragraph.AddHyperlink`), leerlos
(`Paragraph.Hyperlinks`, `Hyperlink`) y editar su texto y destino en el cuerpo, celdas de tabla y
encabezados/pies. Cada enlace usa una relación de hipervínculo de la parte que lo contiene; editar
el destino no cambia otros enlaces que compartían la relación y elimina la relación anterior cuando
queda sin uso. Los enlaces internos a marcadores (`w:anchor`), los atributos de hipervínculo que no
son el destino y el estilo de carácter automático quedan fuera del incremento y se preservan (ver
`docs/adr/0019-external-hyperlinks.md`).

Nota: el 2.10 modela las **propiedades básicas del documento** (`docProps/core.xml`): título, asunto,
autor, palabras clave, descripción, último editor, revisión, fechas de creación/modificación/impresión,
categoría, estado, identificador, idioma y versión, con lectura y edición vivas (`WordDocument.Properties`,
`DocumentProperties`). Las propiedades extendidas (`docProps/app.xml`) y las personalizadas
(`docProps/custom.xml`) se preservan pero no se modelan; las fechas se conservan como instante y el SDK
las devuelve como `DateTime` local (ver `docs/adr/0020-document-properties.md`).

Resultados reales: ver `docs/phase-2-report.md` (incrementos 2.1, 2.2, 2.2b, 2.3, 2.4, 2.5a, 2.5b,
2.5c, 2.6, 2.7, 2.8, 2.9 y 2.10).

Criterio transversal: el texto visible puede repartirse en varios runs; el alcance de la
búsqueda/reemplazo debe definirse expresamente y probarse.

## Fase 3 — Excel avanzado inicial — **Completada**

Cada incremento exige round-trip, verificación con una herramienta independiente y
documentación de limitaciones.

| # | Incremento | Criterio de aceptación | Estado |
| --- | --- | --- | --- |
| 3.1 | Acceso a celdas por índice de fila/columna (1-based) y dirección de la celda (`Cell.Row`/`Column`) | Round-trip por archivo y stream; paquete válido con el SDK; la misma celda se direcciona por A1 y por índice; un índice fuera de rango falla con `ArgumentOutOfRangeException`; leer una celda ausente no la crea | Hecho |
| 3.2 | Rangos (`A1:B2`): direccionamiento, recorrido y lectura/escritura de valores | Alcance y semántica definidos en ADR; round-trip; validación con el SDK; celdas ausentes y tipos mezclados tratados explícitamente | Hecho |
| 3.3 | Fecha/hora y formatos numéricos | Round-trip del serial y del formato; escritura de `DateTime` con formato explícito; lo no modelado se preserva | Hecho |
| 3.4 | Estilos reutilizables (fuente, relleno, bordes, alineación) | Round-trip; validación con el SDK; `cellXfs` reutilizados sin duplicar estilos | Hecho |
| 3.5 | Escritura de fórmulas (separada del cálculo) y conservación del resultado almacenado | La fórmula se escribe como fórmula; el valor en caché se conserva o se declara ausente; no se calcula nada | Hecho |
| 3.6a | Rangos con nombre (`definedNames`) | Round-trip de cada variante soportada (ámbito de libro y de hoja); lo no modelado se conserva | Hecho |
| 3.6b | Validación de datos | Round-trip de las variantes soportadas; lo no modelado se preserva | Hecho |
| 3.6c | Autofiltros | Round-trip; validación con el SDK | Hecho |
| 3.6d | Tablas estructuradas | Round-trip de las variantes soportadas; relaciones preservadas | Hecho |
| 3.7 | Congelar paneles, ocultar filas/columnas, alto/ancho | Round-trip; validación con el SDK | Hecho |
| 3.8 | Configuración de impresión y página | Round-trip; validación con el SDK | Hecho |
| 3.9 | Hipervínculos y notas/comentarios | Round-trip; relaciones preservadas | Hecho |
| 3.10a | Imágenes en la hoja (tipos y propiedades enumerados expresamente) | Round-trip de los tipos soportados; los no soportados se preservan | Hecho |
| 3.10b | Gráficos | Round-trip de los tipos soportados; los no soportados se preservan | Hecho |
| 3.11 | Optimización a tabla de cadenas compartidas | Deduplicación; round-trip; los archivos ajenos con `sharedStrings` se siguen leyendo | Hecho |

Nota: el punto 1 original («Rangos y API por índice de fila/columna») se dividió en **3.1**
(direccionamiento de una celda por índice) y **3.2** (rangos), **en lugar de declarar hecho algo a
medias**. El 3.1 no define qué es un rango; el 3.2 no cambia la dirección de una celda. El 3.3
define los seriales 1900/1904 y la relación entre fecha y formato numérico. El 3.4 separa el formato
numérico (`Cell.NumberFormat`) del estilo visual reutilizable (`Cell.Style`: fuente, relleno sólido,
cuatro bordes y alineación). El 3.5 escribe fórmulas (`Cell.Formula`) sin calcular nada: la fórmula se
guarda como fórmula, el resultado almacenado se conserva al leer y se declara ausente al escribir, y
el libro pide un recálculo al abrir. Ver `docs/adr/0021-excel-cell-indices.md`,
`docs/adr/0022-excel-cell-ranges.md`, `docs/adr/0023-excel-dates-and-number-formats.md`,
`docs/adr/0024-excel-reusable-cell-styles.md` y `docs/adr/0025-excel-formula-writing.md`.

El punto 3.6 original («Rangos con nombre, validación de datos, autofiltros y tablas
estructuradas») se divide en **3.6a** (rangos con nombre), **3.6b** (validación de datos), **3.6c**
(autofiltros) y **3.6d** (tablas estructuradas), para poder declarar hecho cada capacidad por
separado. Ver `docs/adr/0026-excel-named-ranges.md`,
`docs/adr/0027-excel-data-validation.md`, `docs/adr/0028-excel-auto-filters.md` y
`docs/adr/0029-excel-structured-tables.md`.

Resultados reales: ver `docs/phase-3-report.md`.

Criterio: publicar la matriz de funciones de cálculo antes de implementar cualquier motor, y
no afirmar equivalencia con Excel.

## Fase 4 — PDF — **Completada**

Interfaz `IPdfConverter` + adaptador LibreOffice headless, con verificación de ejecutable,
directorios temporales aislados, límites, manejo de errores y exportación desde el estado en
memoria. Las pruebas de integración se omiten con motivo explícito si falta el motor.

| # | Incremento | Criterio de aceptación | Estado |
| --- | --- | --- | --- |
| 4.1 | Contratos de conversión (`OfficeEngine.Pdf`) | El núcleo no depende de ningún motor; opciones, resultado, límites y excepción tipada | Hecho |
| 4.2 | Adaptador LibreOffice headless (`OfficeEngine.Pdf.LibreOffice`) | Motor verificado (versión, licencia, headless); aislamiento por conversión; código de salida y PDF comprobados; pruebas de integración ejecutadas | Hecho |

## Fase 5 — Robustez — **Completada**

Archivos corruptos y maliciosos, profundidad de XML, entidades externas, grandes documentos,
regresión, pruebas de interoperabilidad y benchmarks reproducibles. Revisión de conservación de
partes no soportadas.

| # | Incremento | Criterio de aceptación | Estado |
| --- | --- | --- | --- |
| 5.1 | Carga de OOXML hostil (profundidad de XML, DTD, entidades externas) | Un XML anidado en exceso falla con un error tipado en lugar de terminar el proceso; los DTD se rechazan de forma explícita; el límite es público y configurable | Hecho |
| 5.2 | Corpus de archivos corruptos y errores tipados | Cada mutación de un paquete válido falla con un error tipado o se abre; ninguna excepción del SDK o del lector de XML llega a la API pública | Hecho |
| 5.3 | Documentos grandes, modelo de almacenamiento y `async` | Se decide y se documenta el modelo de memoria y de cancelación; se miden los límites reales | Hecho |
| 5.4 | Benchmarks reproducibles y regresión | Hay una base de rendimiento reproducible en CI y una regresión de round-trip sobre documentos generados | Hecho |

## Fase 6 — DOC/XLS y funciones avanzadas adicionales — **Completada**

Conversión externa opcional con pérdidas documentadas; funciones complejas solo con pruebas
reproducibles y matriz actualizada.

| # | Incremento | Criterio de aceptación | Estado |
| --- | --- | --- | --- |
| 6.1 | Conversión de los formatos heredados (`.doc`, `.xls`) con un motor externo | Convierte, no lee: el resultado se abre con la API moderna; las pérdidas están documentadas; verificación diferencial por viaje de ida y vuelta | Hecho |
| 6.2 | Motor de fórmulas propio (`OfficeEngine.Formulas`) | Matriz de funciones publicada; los valores se comparan con un motor externo; los errores de Excel se distinguen | Hecho |

## Fase 7 — Publicación

NuGet, versionado semántico, changelog, guía de inicio, referencia de API, ejemplos
(`OfficeEngine.Samples`), SBOM y avisos de licencias de terceros.
