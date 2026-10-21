# Fase 3 — Informe de resultados reales (incrementos 3.1 a 3.11)

Fecha: 2026-10-20.
Entorno: Windows 11 Enterprise LTSC (10.0.26100, x64), .NET SDK 8.0.426.
Comandos ejecutados y resultados **observados** (no estimados).

> La Fase 3 empieza con el acceso a celdas por índice de fila/columna, sigue con los rangos de
> celdas, las fechas y los formatos numéricos, continúa con los estilos reutilizables, escribe
> fórmulas sin calcularlas y termina esta tanda con los rangos con nombre, la validación de datos,
> los autofiltros y las tablas estructuradas. El motor de cálculo, la configuración de página, las
> imágenes y los gráficos siguen sin implementarse; consulta `docs/compatibility-matrix.md` para el
> estado real, capacidad por capacidad.

## Alcance del incremento 3.1

Direccionar una celda por **índice 1-based de fila y columna**, sin cambiar la representación A1 que
se guarda en el paquete. Incluye:

- `Worksheet.GetCell(int row, int column)` y el indexador `Worksheet[int row, int column]`, con la
  fila primero y la columna después (`sheet[7, 2]` es `B7`).
- `Cell.Row` y `Cell.Column`, que informan de la posición 1-based de la celda junto a
  `Cell.Reference` (A1).
- Límites del formato validados antes de tocar el paquete: fila `1..1048576` y columna
  `1..16384`; un índice fuera de rango lanza `ArgumentOutOfRangeException` con el parámetro
  rechazado.
- La misma semántica que la referencia A1: el handle es una vista viva, leer una celda ausente no la
  crea y escribir en ella crea la fila y la celda en orden de columna.

Los **rangos** (`A1:B2`) quedaron para el incremento 3.2, que se documenta más abajo. La hoja de ruta
y el ADR 0021 documentan la división.

## Comandos y resultados del 3.1

| Comando | Resultado |
| --- | --- |
| `dotnet test OfficeEngine.sln -c Release` (antes del cambio) | **632 superadas, 0 con error, 0 omitidas** (línea base). |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **648 superadas, 0 con error, 0 omitidas** (648 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3` | **648 superadas**; cobertura generada en `artifacts/testresults/phase3/**/coverage.cobertura.xml`. |

Salida textual del resumen de las pruebas (la duración varía entre ejecuciones):

```
Correctas! - Con error:     0, Superado:   648, Omitido:     0, Total:   648
```

Las **632 pruebas anteriores siguen pasando** sin cambios.

## Cobertura del 3.1 (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/Cell.cs` | 62 / 75 | 82,7 % | 70,8 % |
| `Excel/Worksheet.cs` | 64 / 66 | 97,0 % | 75,9 % |
| **Total de la solución** | — | 94,8 % | 88,6 % |

Los miembros nuevos (`Cell.Row`/`Column`, `Worksheet.GetCell(int, int)` y el indexador) están
cubiertos. Las líneas sin cubrir de `Cell.cs` son rutas preexistentes de lectura de cadenas
compartidas mal formadas, la lectura de un tipo `String`/`Error` almacenado, la conversión de tipos
numéricos distintos de `double` y `ToString`; las de `Worksheet.cs` son el `ToString` y el error de
una parte de hoja ilegible, también preexistentes.

## Pruebas nuevas del 3.1 (16 casos)

| Archivo | Foco |
| --- | --- |
| `Excel/ExcelCellIndexTests.cs` (15) | Dirección 1-based (`[1,1]` = `A1`); `Row`/`Column` de un handle A1; misma celda por A1 y por índice en ambos sentidos; round-trip por stream; última fila y columna (`XFD1048576`); seis índices fuera de rango; nombre del parámetro rechazado (`row`/`column`); lectura de una celda ausente sin crearla; orden de columnas al escribir C1, A1 y B1; libro liberado. |
| `Excel/ExcelInteropTests.cs` (+1) | Paquete válido con el SDK y lectura de una celda escrita por índice (`[2,2]` = `B2`) con `OpenXmlValidator` y `SpreadsheetDocument`. |

## Verificación independiente del 3.1 (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` del paquete producido al escribir por índice
   (`ProducedPackage_WithCellWrittenByIndex_IsValidAndReadableByTheSdk`).
2. **Lectura con el SDK** de la celda `B2` escrita por índice, comprobando que el paquete guarda la
   referencia A1 y no una representación distinta.
3. **Lectura cruda** de `xl/worksheets/sheet1.xml` para comprobar que leer una celda ausente no crea
   `Z99` y que las celdas escritas por índice se insertan en orden de columna.

## Hallazgos del 3.1 durante las pruebas y acciones tomadas

1. **El validador del SDK no puede comprobar los límites de un índice.** El índice no llega al
   paquete: se convierte en una referencia A1. Por eso los límites (`1048576` filas y `16384`
   columnas) se validan en la frontera pública, antes de crear el handle, y se prueban con seis
   combinaciones fuera de rango.
2. **El nombre del parámetro importa.** Un índice inválido debe señalar `row` o `column`, no un
   mensaje genérico, porque el llamante suele recorrer filas y columnas en bucles y necesita saber
   cuál corregir. Hay prueba dedicada (`CellIndexer_OutOfRange_NamesTheRejectedIndex`).
3. **La celda ya conocía sus índices, pero no los exponía.** `Cell` guardaba fila y columna para
   localizar el elemento; exponerlos evita volver a analizar `Reference` en un bucle y no cambia el
   estado del paquete.
4. **Leer una celda ausente sigue siendo una operación de solo lectura.** El camino por índice usa
   `Worksheet.FindCell`, que además conserva la comprobación existente de una fila sin `@r`: un
   archivo malformado falla con error tipado en vez de devolver `null` por una celda que no se puede
   localizar con fiabilidad.

## Qué **no** se ha implementado en 3.1 (y por tanto no se afirma)

- **Rangos** (`A1:B2`): no formaban parte del 3.1; se entregaron en el 3.2 (ver la sección
  siguiente).
- **Referencias absolutas o mixtas** (`$A$1`, `A$1`): no se aceptan en `GetCell(string)` ni se emiten.
- **Referencias 3D** entre hojas (`Hoja1!A1`) y nombres con espacio o comillas.
- **Borrar o mover celdas**, filas o columnas.
- **Estilos, formatos numéricos y fechas al escribir**: quedaron fuera del 3.1; las fechas y los
  formatos numéricos se entregaron en el 3.3.
- **Fórmulas y cálculo**: sin cambios respecto a la Fase 2.
- La decisión completa está en `docs/adr/0021-excel-cell-indices.md`.

## Incremento 3.2 — Rangos de celdas

### Alcance del incremento 3.2

Direccionar una zona rectangular por sus dos esquinas, recorrerla y leer o escribir sus valores sin
cambiar la representación A1 del paquete. Incluye:

- `Worksheet.GetRange(string)` con la forma `B2:D5` (dos esquinas obligatorias, minúsculas
  normalizadas, esquinas invertidas rechazadas) y
  `Worksheet.GetRange(int firstRow, int firstColumn, int lastRow, int lastColumn)`.
- `CellRange` con `Reference`, las cuatro esquinas, `RowCount` y `ColumnCount`.
- `Cells`, una secuencia perezosa en orden de fila que lee el libro vivo y no crea celdas ausentes.
- `GetValues()` y `SetValues(object?[,])`: matriz de `RowCount × ColumnCount` con índices 0-based;
  `null` para una celda ausente; escritura con dimensiones exactas, validación previa de todos los
  valores y sin escritura parcial si uno no es soportado.
- `null` limpia una celda existente y no crea una ausente. Para que la semántica sea una sola,
  `Cell.Value = null` también deja de crear una celda vacía (cambio de comportamiento documentado en
  el ADR 0022).

La decisión completa está en `docs/adr/0022-excel-cell-ranges.md`.

### Comandos y resultados del 3.2

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **679 superadas, 0 con error, 0 omitidas** (679 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc32b` | **679 superadas**; cobertura generada en `artifacts/testresults/phase3/inc32b/**/coverage.cobertura.xml`. |

Salida textual del resumen (la duración varía):

```
Correctas! - Con error:     0, Superado:   679, Omitido:     0, Total:   679
```

Las **648 pruebas anteriores siguen pasando** sin cambios; las 31 nuevas son 30 unitarias y 1 de
interoperabilidad.

### Cobertura del 3.2 (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/CellRange.cs` | 49 / 49 | 100 % | 100 % |
| `Excel/Cell.cs` | 71 / 84 | 84,5 % | 74,4 % |
| `Excel/Worksheet.cs` | 87 / 89 | 97,8 % | 75,0 % |
| `Excel/CellReference.cs` | 53 / 54 | 98,1 % | 97,6 % |
| **Total de la solución** | — | 95,0 % | 88,9 % |

`CellRange` está cubierto al 100 %. Las líneas sin cubrir de `Cell.cs` son las mismas rutas
preexistentes de lectura (cadenas compartidas mal formadas, `String`/`Error`, conversiones numéricas
distintas de `double` y `ToString`); las de `Worksheet.cs` son `ToString` y el error de una parte de
hoja ilegible; la de `CellReference.cs` es una rama de rechazo de referencia vacía.

### Pruebas nuevas del incremento (31 casos)

| Archivo | Foco |
| --- | --- |
| `Excel/ExcelCellRangeTests.cs` (29) | Sintaxis y normalización; ambas sobrecargas; esquinas invertidas y fuera de rango; orden de fila; vista viva; lectura de celdas ausentes sin crearlas; matriz de valores en ambos sentidos; dimensiones exactas; tipo no soportado sin escritura parcial; `null` limpia sin crear; round-trip por stream; libro liberado. |
| `Excel/ExcelWorkbookTests.cs` (+1) | `Cell.Value = null` sobre una celda ausente no la crea (cambio de comportamiento del ADR 0022). |
| `Excel/ExcelInteropTests.cs` (+1) | Paquete válido con el SDK y lectura de una matriz escrita con `SetValues` (`A1:B2`), incluida una celda `null` que no se materializa. |

### Verificación independiente del 3.2 (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` del paquete producido al escribir un rango
   (`ProducedPackage_WithRangeValues_IsValidAndReadableByTheSdk`).
2. **Lectura con el SDK** de las celdas `A1`, `B1` y `A2` escritas con `SetValues`, y comprobación de
   que `B2` (un `null`) no existe en el paquete.
3. **Lectura cruda** de `xl/worksheets/sheet1.xml` para comprobar que leer un rango ausente no crea
   `Z99`/`AA100` y que escribir `null` no crea `A2`/`B2` cuando faltan.
4. **Round-trip por stream**: los valores de `B2:C3` se recuperan con `GetValues` tras guardar y
   reabrir.

### Hallazgos del 3.2 durante las pruebas y acciones tomadas

1. **Una sola celda no es un rango.** `GetRange("A1")` se rechaza para no duplicar la API de celda ni
   inventar una referencia canónica `A1:A1`; el error indica que se use el indexador de celda.
2. **Las esquinas invertidas se rechazan, no se reordenan.** Reordenar en silencio cambiaría la zona
   pedida sin avisar; el llamante debe construir el rectángulo en orden.
3. **`null` no debe materializar la matriz.** La versión anterior de `Cell.Value = null` creaba una
   celda vacía. Se cambió a «limpiar si existe, no crear si falta» y se añadió prueba; la prueba
   anterior de `null` sigue pasando. El cambio evita que un rango de `null` infle el paquete.
4. **Validar antes de escribir evita un rango a medias.** `SetValues` recorre la matriz una vez para
   validar los tipos y solo después escribe; un `Guid` en la segunda celda deja la primera intacta.
5. **La matriz es 0-based aunque la hoja sea 1-based.** Es la convención de los arrays de .NET y está
   documentada; fingir índices 1-based en un array requeriría una abstracción que nadie pidió.
6. **`Cells` es perezoso.** Materializar una lista de un rango grande podía agotar la memoria; la
   secuencia se genera al enumerar. `GetValues`/`SetValues` siguen siendo explícitamente masivos y
   celda a celda; su coste superlineal queda anotado como `shortcut:` para la Fase 5.

### Qué **no** se ha implementado en 3.2 (y por tanto no se afirma)

- Rangos con nombre, referencias absolutas o mixtas (`$A$1`), columnas o filas completas (`A:A`,
  `1:1`), referencias 3D, uniones e intersecciones.
- Copiar, mover o borrar un rango como operación atómica; desplazarlo o redimensionarlo.
- Estilos, formatos numéricos y fechas: quedaron fuera del 3.2; las fechas y los formatos numéricos
  se entregaron en el 3.3 y la escritura de fórmulas en el 3.5.
- Expansión o contracción de celdas combinadas: el rango las recorre celda a celda.
- Optimización del acceso masivo: se medirá en la Fase 5 antes de cambiarla.

## Incremento 3.3 — Fechas, horas y formatos numéricos

### Alcance del incremento 3.3

Leer y escribir una fecha como el serial de Excel y el formato que la muestra como fecha, sin
perder los estilos que esta versión no modela. Incluye:

- `Cell.NumberFormat`: lee el código efectivo (`General` si no hay formato explícito ni heredado,
  `null` para un formato integrado no modelado) y escribe un formato nuevo o vuelve a `General` con
  `null`.
- `Cell.SetDateTime(DateTime, string)`: escribe el serial y un formato de fecha/hora explícito en una
  sola llamada. `Cell.Value = DateTime` solo se acepta si la celda ya tiene un formato de fecha/hora;
  no se adivina ninguno.
- Seriales de Excel en los sistemas 1900 y 1904. En el 1900 se respeta el error histórico del
  `1900-02-29`: los seriales `[60, 61)` y los menores que 1 se leen como número, no como fecha.
- Detección de fecha/hora en el código de formato: tokens `y/m/d/h/s/b/g/e`, `A/P` y `AM/PM`, sin
  confundir la notación científica `E+`/`E-` ni los literales entre comillas.
- Formatos personalizados con ids a partir de 164, reutilizados por código; creación perezosa de un
  `xl/styles.xml` mínimo y preservación del stylesheet existente.
- `CellRange.GetValues` devuelve `DateTime` para celdas con formato de fecha/hora.
  `CellRange.SetValues` rechaza `DateTime` porque cada celda necesita su propio formato explícito.

La decisión completa está en `docs/adr/0023-excel-dates-and-number-formats.md`.

### Comandos y resultados del 3.3

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **736 superadas, 0 con error, 0 omitidas** (736 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc33` | **736 superadas**; cobertura generada en `artifacts/testresults/phase3/inc33/**/coverage.cobertura.xml`. |

Salida textual del resumen (la duración varía):

```
Correctas! - Con error:     0, Superado:   736, Omitido:     0, Total:   736
```

Las **679 pruebas anteriores siguen pasando**; las 57 nuevas son 54 unitarias/de round-trip en
`ExcelDateAndNumberFormatTests`, 2 en `ExcelCellRangeTests` y 1 de interoperabilidad en
`ExcelInteropTests`. La prueba existente de tipos no soportados se actualizó: `DateTime` ya no es un
tipo rechazado y solo `Guid` sigue siéndolo.

### Cobertura del 3.3 (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelStyles.cs` (`ExcelStyles`) | 195 / 233 | 83,7 % | 66,4 % |
| `Excel/ExcelStyles.cs` (`NumberFormatCode`) | 40 / 40 | 100 % | 91,8 % |
| `Excel/Cell.cs` | 145 / 163 | 89,0 % | 81,5 % |
| `Excel/CellRange.cs` | 49 / 49 | 100 % | 100 % |
| `Excel/Worksheet.cs` | 87 / 89 | 97,8 % | 75,0 % |
| `Excel/ExcelWorkbook.cs` | 101 / 122 | 82,8 % | 69,3 % |
| **Total de la solución** | 3366 / 3570 | 94,3 % | 87,7 % |

`NumberFormatCode` está cubierto al 100 % de líneas. Las líneas sin cubrir de `ExcelStyles.cs` son
principalmente rutas de error de un stylesheet malformado, el agotamiento de ids de formato y la
creación de `cellXfs` cuando el stylesheet existente no los tiene. Las de `Cell.cs` son las mismas
rutas preexistentes de lectura (cadenas compartidas mal formadas, `String`/`Error`, conversiones
numéricas distintas de `double` y `ToString`) y los seriales no finitos o fuera de rango.

### Pruebas nuevas del 3.3 (57 casos)

| Archivo | Foco |
| --- | --- |
| `Excel/ExcelDateAndNumberFormatTests.cs` (54) | `NumberFormat` en celda ausente, existente y liberada; `General`/`null`; formatos inválidos; `SetDateTime` con fecha, hora y formatos explícitos; rechazo de formatos sin tokens; sistema 1900 (`1`, `59`, `61`) y 1904; seriales `0`, `0,5`, `60` y `60,5` devueltos como número; formatos integrados y personalizados; notación científica y literales; preservación de fuente, relleno, borde y alineación al cambiar o resetear el formato; formato integrado de fecha dependiente de la configuración regional; celda sin `s` cuyo formato está en el `cellXf` 0. |
| `Excel/ExcelCellRangeTests.cs` (+2) | `GetValues` devuelve `DateTime`; `SetValues` rechaza `DateTime` sin escritura parcial. |
| `Excel/ExcelInteropTests.cs` (+1) | Paquete válido con el SDK, serial y número leídos por el SDK, y códigos de formato leídos desde `cellXfs`/`numFmts`. |

### Verificación independiente del 3.3 (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` del paquete producido con una fecha y un formato
   numérico (`ProducedPackage_WithDateAndNumberFormats_IsValidAndReadableByTheSdk`).
2. **Lectura con el SDK** del serial `45322` de una fecha, del número `1234.5` y de los códigos
   `yyyy-mm-dd` y `#,##0.00` desde `cellXfs`/`numFmts`.
3. **Fixtures construidos con el SDK** (`OoxmlValidation.CreateXlsxWithFormattedCell`), no con la
   librería bajo prueba: formato integrado 14, formato personalizado 164, sistema 1904, serial 60,
   formato científico y un estilo con fuente, relleno, borde y alineación.
4. **Lectura del estilo resultante con el SDK** (`ReadXlsxCellStyle`): el cambio de formato conserva
   los ids de fuente, relleno y borde, el XML conserva la alineación y el reseteo a `General` también
   los conserva.
5. **Round-trip por stream**: la fecha y el formato se recuperan tras guardar y reabrir.

### Hallazgos del 3.3 durante las pruebas y acciones tomadas

1. **Una fecha es un serial más un formato.** Sin un formato de fecha/hora no hay forma de distinguir
   una fecha de un número. Por eso `Cell.Value = DateTime` exige un formato ya presente y
   `SetDateTime` recibe el formato de forma explícita; no se inventa `yyyy-mm-dd`.
2. **El error del 29 de febrero de 1900 no se puede corregir sin mentir.** El serial 60 no
   corresponde a ninguna fecha real; se devuelve como número y se conserva el formato. Lo mismo se
   hace con los seriales menores que 1.
3. **El sistema 1904 no es una curiosidad.** Un libro abierto con `date1904` se lee y se escribe con
   ese origen; la fecha `2024-01-01` es el serial `43830`, no `45292`.
4. **Un stylesheet existente no se reescribe.** El primer formato crea un stylesheet mínimo; en un
   libro abierto solo se añaden entradas nuevas y se actualizan los contadores. Un cambio de formato
   clona el `cellXf` actual para conservar fuente, relleno, borde y alineación.
5. **La notación científica no es una fecha.** `0.00E+00` contiene una `E`, pero no es el token de
   era; el detector la ignora y la celda se lee como número. Los literales entre comillas y los
   corchetes de color tampoco cuentan.
6. **Una matriz no puede llevar formatos por celda.** `CellRange.SetValues` rechaza `DateTime` en la
   validación previa, de modo que el rango queda intacto; la fecha se escribe celda a celda con
   `SetDateTime`.

### Qué **no** se ha implementado en 3.3 (y por tanto no se afirma)

- Estilos reutilizables (fuente, relleno, bordes, alineación) como API pública: quedan para el 3.4.
- Escritura de fórmulas y cálculo: la escritura llegó en el 3.5; el cálculo sigue siendo un subsistema
  separado.
- Un modelo de solo fecha o solo hora, o la escritura de una hora sin fecha. Un `DateTime` siempre
  tiene componente de fecha; los seriales menores que 1 se leen como número.
- La exposición del código de los formatos integrados dependientes de la configuración regional
  (ids 27–36 y 50–58): se reconocen como fecha/hora, pero `NumberFormat` devuelve `null`.
- La validación semántica completa del lenguaje de formatos de Excel y la conversión de
  `DateTimeKind`; los componentes de fecha y hora se guardan tal cual.
- La detección de fecha/hora es heurística: no cubre todo el lenguaje de formatos y un marcador
  AM/PM local que no use caracteres ASCII puede no detectarse; el valor se lee como número y el
  formato se preserva.
- Cambiar el sistema de fecha de un libro (`date1904`); solo se respeta el que ya tiene.
- Escribir `DateTime` con `CellRange.SetValues` o asignar un formato automático al escribir una
  fecha.

## Incremento 3.4 — Estilos reutilizables de celda

### Alcance del incremento 3.4

Modelar un estilo visual reutilizable sin mezclarlo con el formato numérico de 3.3. Incluye:

- `Cell.Style` con un `CellStyle` inmutable: fuente (nombre, tamaño, negrita, cursiva, subrayado y
  color RGB), relleno sólido, cuatro bordes por arista y alineación horizontal/vertical con ajuste
  de texto.
- `CellBorder`, `CellBorders`, `CellBorderStyle`, `CellUnderlineStyle`, `CellHorizontalAlignment` y
  `CellVerticalAlignment` como tipos públicos.
- Semántica tri-state: `null` vuelve al valor por defecto del libro; un `CellStyle` vacío resetea las
  propiedades modeladas y conserva el formato numérico.
- Creación perezosa del stylesheet y reutilización por igualdad exacta de XML de `font`, `fill`,
  `border` y `cellXfs`. Aplicar el mismo estilo a varias celdas no duplica estilos.
- Preservación del formato numérico, del resto del `xf` y de las propiedades no modeladas de fuente,
  borde y alineación. El relleno no sólido es la excepción documentada.

La decisión completa está en `docs/adr/0024-excel-reusable-cell-styles.md`.

### Comandos y resultados del 3.4

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **764 superadas, 0 con error, 0 omitidas** (764 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc34` | **764 superadas**; cobertura generada en `artifacts/testresults/phase3/inc34/**/coverage.cobertura.xml`. |

Salida textual del resumen (la duración varía):

```
Correctas! - Con error:     0, Superado:   764, Omitido:     0, Total:   764
```

Las **736 pruebas anteriores siguen pasando**; las 28 nuevas son 27 en `ExcelCellStyleTests` y 1 de
interoperabilidad en `ExcelInteropTests`.

### Cobertura del 3.4 (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelStyles.cs` (`ExcelStyles`) | 581 / 668 | 87,0 % | 72,4 % |
| `Excel/ExcelStyles.cs` (`NumberFormatCode`) | 40 / 40 | 100 % | 91,8 % |
| `Excel/Cell.cs` | 157 / 175 | 89,7 % | 82,4 % |
| `Excel/CellStyle.cs` | 11 / 11 | 100 % | 100 % |
| `Excel/CellBorder.cs` | 9 / 12 | 75,0 % | 60,0 % |
| `Excel/CellBorders.cs` | 12 / 15 | 80,0 % | 60,0 % |
| `Excel/CellRange.cs` | 49 / 49 | 100 % | 100 % |
| **Total de la solución** | 3796 / 4055 | 93,6 % | 85,3 % |

Las líneas sin cubrir de `ExcelStyles.cs` son rutas de error de un stylesheet malformado, componentes
o `xf` inexistentes, valores de borde/alineación no soportados y algunas ramas de preservación. Las
de `CellBorder.cs` y `CellBorders.cs` son principalmente los operadores de igualdad y sus rutas de
`GetHashCode`/`Equals(object)`. La cobertura baja respecto al 3.3 porque el incremento añade código
nuevo de validación y preservación; no hay regresiones en las pruebas existentes.

### Pruebas nuevas del 3.4 (28 casos)

| Archivo | Foco |
| --- | --- |
| `Excel/ExcelCellStyleTests.cs` (27) | Celda ausente sin creación; estilo completo con round-trip; reutilización de componentes y `cellXfs`; estilos distintos; formato numérico y fecha conservados; reseteo con estilo vacío; celda creada sin valor; validación de nombre, tamaño, enums y `null`; los 14 bordes, 5 subrayados y todas las alineaciones; lectura de un fixture ajeno; preservación de `strike`, `diagonal` e `indent`; libro liberado. |
| `Excel/ExcelInteropTests.cs` (+1) | Paquete válido con el SDK, misma entrada de `cellXfs` para dos celdas y lectura de fuente, color, relleno, borde y alineación. |

### Verificación independiente del 3.4 (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de paquetes con fuente, relleno, bordes y
   alineación (`ProducedPackage_WithCellStyles_IsValidAndReadableByTheSdk` y los round-trips de
   `ExcelCellStyleTests`).
2. **Lectura con el SDK** de la misma entrada de `cellXfs` para dos celdas y de los valores de fuente,
   color, relleno, borde y alineación.
3. **Fixtures construidos con el SDK** (`OoxmlValidation.CreateXlsxWithFormattedCell`), no con la
   librería bajo prueba, incluyendo un estilo con `strike`, borde diagonal e `indent` no modelados.
4. **Lectura del estilo resultante con el SDK** (`ReadXlsxCellStyleParts` y `ReadXlsxStyleCounts`):
   recuentos de `fonts`/`fills`/`borders`/`cellXfs` y preservación de las propiedades no modeladas.
5. **Round-trip por stream** de los valores, del formato numérico y de todos los estilos modelados.

### Hallazgos del 3.4 durante las pruebas y acciones tomadas

1. **Un estilo no es solo un `xf`.** Para reutilizar de verdad hay que deduplicar también `font`,
   `fill` y `border`; de lo contrario, tres celdas con el mismo estilo acaban con tres fuentes
   idénticas. La búsqueda por `OuterXml` resuelve ambos niveles.
2. **El formato numérico no pertenece a `CellStyle`.** Incluirlo habría obligado a inventar un valor
   para los formatos integrados no modelados y habría perdido el formato al aplicar un estilo. Se
   mantiene en `Cell.NumberFormat` y el `xf` se clona para conservarlo.
3. **Aplicar un estilo no debe borrar lo que la API no entiende.** La fuente, el borde y la alineación
   se clonan y solo se reescriben sus hijos/atributos modelados; así sobreviven `strike`, `diagonal`,
   `indent` y el resto. El relleno no sólido no se puede conservar porque el modelo solo expresa
   relleno sólido, y se documenta.
4. **Los colores de SpreadsheetML son ARGB.** Escribir seis dígitos producía un paquete que el
   validador aceptaba pero que no era el formato habitual de Excel; se escribe `FF` + RGB y se acepta
   ARGB al leer, ignorando el alfa.
5. **Un estilo vacío en una celda ausente no debe crearla.** La comprobación de estilo por defecto se
   hace antes de crear la celda; una celda existente sí se resetea.

### Qué **no** se ha implementado en 3.4 (y por tanto no se afirma)

- Rellenos con patrón, degradados o colores de tema; el relleno sólido es el único modelado.
- Protección de celda, rotación de texto, sangría, `shrinkToFit`, orden de lectura, justificación de
  la última línea y el resto de atributos de alineación.
- Estilos con nombre, estilos de tabla, formato condicional, `dxf` y la resolución de la herencia
  `cellStyleXfs`/`xfId`.
- Aplicar un estilo a un rango de una sola vez; se recorre con `CellRange.Cells`.
- Conservar el canal alfa o el marcado de color de tema; el color se reduce a RGB.

## Incremento 3.5 — Escritura de fórmulas (sin cálculo)

### Alcance del incremento 3.5

Escribir una fórmula como fórmula, sin mezclarla con el cálculo y sin perder el resultado
almacenado. Incluye:

- `Cell.Formula`: getter con el `=` inicial (`=A1*B1`) y setter que acepta la fórmula con o sin él y
  guarda una fórmula normal (`<f>` sin atributos de tipo).
- Semántica del resultado almacenado: al escribir una fórmula se elimina el `<v>` (resultado
  ausente); al leer se devuelve el que había; al abrir y guardar sin tocar la celda se conservan
  ambos. Escribir un valor literal sobre una celda con fórmula elimina la fórmula.
- `Cell.Formula = null` elimina fórmula y resultado; una celda ausente sigue ausente.
- Validación antes de tocar el paquete: fórmula vacía, de más de 8192 caracteres o con caracteres no
  válidos en XML.
- Petición de recálculo al abrir: `calcPr/@fullCalcOnLoad="1"`, conservando el resto de atributos del
  `calcPr` e insertándolo en el orden del esquema si no existía.
- Ningún cálculo, ningún grafo de dependencias y ningún `eval`.

La decisión completa está en `docs/adr/0025-excel-formula-writing.md`.

### Comandos y resultados del 3.5

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **776 superadas, 0 con error, 0 omitidas** (776 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc35` | **776 superadas**; cobertura generada en `artifacts/testresults/phase3/inc35/**/coverage.cobertura.xml`. |

Salida textual del resumen (la duración varía):

```
Correctas! - Con error:     0, Superado:   776, Omitido:     0, Total:   776
```

Las **764 pruebas anteriores siguen pasando**; las 12 nuevas están en `ExcelFormulaTests` y se
añadieron aserciones a `ExcelInteropTests.OpenedWorkbook_WithFormula_ReturnsCachedValueAndPreservesFormula`
y a `ReadmeExamplesTests.Excel_CreateWriteCellsAndSave`.

### Cobertura del 3.5 (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelFormula.cs` | 41 / 41 | 100 % | 88,9 % |
| `Excel/Cell.cs` | 176 / 194 | 90,7 % | 82,9 % |
| **Total de la solución** | 3856 / 4115 | 93,7 % | 85,3 % |

Las líneas sin cubrir de `Cell.cs` son anteriores a este incremento (`ToString`, rutas de error de
cadenas compartidas y de serial fuera de rango). Las ramas sin cubrir de `ExcelFormula.cs` son la
lectura de una fórmula ajena que ya empieza por `=` y la ruta de libro sin parte de libro legible.

### Pruebas nuevas del 3.5 (12 casos)

| Archivo | Foco |
| --- | --- |
| `Excel/ExcelFormulaTests.cs` (12) | Escritura y round-trip sin resultado almacenado; normalización del `=` inicial; ausencia de cálculo; `Formula = null` borra fórmula y caché; escribir fórmula borra la caché obsoleta; escribir un valor borra la fórmula; `Value = null` borra ambas; validación (vacía, `=`, 8193 caracteres, carácter no válido en XML) sin crear ni modificar la celda; preservación de formato numérico y estilo; petición de recálculo con `definedNames` y con un elemento posterior (`fileRecoveryPr`); libro liberado. |

### Verificación independiente del 3.5 (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con fórmula, con y sin resultado
   almacenado, y de los paquetes con `definedNames`.
2. **Lectura con el SDK** de la fórmula (`<f>`) y del resultado almacenado (`<v>`) de la celda, y de
   `calcPr/@fullCalcOnLoad`, con `OoxmlValidation.ReadXlsxCellFormula` y
   `ReadXlsxFullCalculationOnLoad`.
3. **Fixtures construidos con el SDK** (`OoxmlValidation.CreateXlsxWithFormula`), no con la librería
   bajo prueba, incluida una variante con `definedNames` y `fileRecoveryPr` para comprobar el orden
   de los elementos del libro.
4. **Round-trip por stream** de la fórmula, del resultado ausente y del resultado almacenado de un
   archivo ajeno.

### Hallazgos del 3.5 durante las pruebas y acciones tomadas

1. **Escribir un valor dejaba la fórmula puesta.** El modelo anterior borraba el valor pero no el
   `<f>`, de modo que una celda podía quedar con una fórmula antigua y un valor nuevo, que es
   contradictorio. Ahora `Cell.Value` (y `Value = null`) eliminan la fórmula, y la caché se borra
   siempre que se toca el contenido de una celda.
2. **El resultado almacenado no se puede inventar.** Al escribir una fórmula se declara ausente: se
   borra el `<v>` y el tipo de dato. Conservar el valor anterior habría creado un resultado que no
   corresponde a la fórmula nueva.
3. **Pedir el recálculo obliga a conocer el orden del libro.** `calcPr` va después de `definedNames`
   y antes de `fileRecoveryPr`; insertarlo al final rompía el esquema en los libros que ya tenían
   esos elementos. La inserción respeta el orden y se comprueba con el validador del SDK.
4. **El `=` inicial es una convención de la API, no del formato.** En el paquete la fórmula se
   guarda sin `=`, como manda SpreadsheetML; la API lo acepta y lo devuelve.
5. **Nada de `eval`.** La fórmula se guarda como texto y no se valida su semántica: `=1+` se escribe
   tal cual. El motor que abre el archivo es quien la evalúa.

### Qué **no** se ha implementado en 3.5 (y por tanto no se afirma)

- Cálculo de fórmulas: no hay motor, ni matriz de funciones, ni grafo de dependencias, ni recálculo
  propio. El recálculo se delega en el motor que abre el archivo.
- Fórmulas compartidas, de matriz y de tabla: se conservan al leer y guardar, pero no se escriben ni
  se expanden; una compartida sin texto propio lee `null`.
- Validación semántica de la expresión, referencias externas, vínculos entre libros, funciones
  volátiles y matrices dinámicas.
- Escribir fórmulas en bloque con `CellRange.SetValues`: el rango escribe valores literales y una
  cadena que empieza por `=` es texto.
- Una API para fijar el resultado almacenado de una fórmula: el resultado lo escribe el motor que
  calcula.

## Incremento 3.6a — Rangos con nombre

### Alcance del incremento 3.6a

Modelar los nombres definidos del libro (`xl/workbook.xml`), incluidos los que esta versión no
interpreta. Incluye:

- `ExcelWorkbook.DefinedNames`: catálogo de solo lectura con los nombres del libro.
- `ExcelWorkbook.FindDefinedName(name)` y `ExcelWorkbook.RemoveNamedRange(name)`.
- `Worksheet.AddNamedRange(name, rangeReference)` (ámbito de libro) y
  `Worksheet.AddLocalNamedRange(name, rangeReference)` (ámbito de hoja), más
  `Worksheet.RemoveLocalNamedRange(name)`.
- `ExcelDefinedName` con `Name`, `RefersTo`, `Scope`, `IsRange` y `Range`.
- Validación del texto del nombre y del rango antes de tocar el paquete.
- Ajuste del ámbito posicional de los nombres locales al eliminar una hoja.

La decisión completa está en `docs/adr/0026-excel-named-ranges.md`.

### Comandos y resultados del 3.6a

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **790 superadas, 0 con error, 0 omitidas** (790 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc36a` | cobertura generada en `artifacts/testresults/phase3/inc36a/**/coverage.cobertura.xml`. |

Las **776 pruebas anteriores siguen pasando**; las 14 nuevas están en `ExcelDefinedNameTests`.

### Cobertura del 3.6a (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelDefinedName.cs` | 12 / 13 | 92,3 % | 50,0 % |
| `Excel/ExcelDefinedNames.cs` | 109 / 126 | 86,5 % | 76,2 % |
| `Excel/Worksheet.cs` | 132 / 137 | 96,4 % | 77,1 % |
| `Excel/ExcelWorkbook.cs` | 163 / 188 | 86,7 % | 73,3 % |
| **Total de la solución** | 4086 / 4370 | 93,5 % | 84,7 % |

Las líneas sin cubrir son rutas de error: referencias mal formadas que llegan desde un archivo ajeno
y casos de análisis que no se pueden producir con la API pública.

### Verificación independiente del 3.6a (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con nombres de libro y de hoja.
2. **Lectura con el SDK** de `<definedNames>` (`OoxmlValidation.ReadXlsxDefinedNames`): texto,
   nombre y `localSheetId` de cada nombre.
3. **Fixtures construidos con el SDK** (`OoxmlValidation.CreateXlsxWithDefinedNames`) con nombres de
   libro, de hoja, constantes y nombres integrados.
4. **Round-trip por stream** del catálogo y de los nombres creados.

### Hallazgos del 3.6a durante las pruebas y acciones tomadas

1. **El ámbito de un nombre local es posicional.** `localSheetId` es el índice de la hoja en la
   lista, no un identificador estable. Al eliminar una hoja, los nombres de esa hoja se eliminan y
   los de las hojas posteriores se desplazan, o el nombre acabaría apuntando a otra hoja.
2. **Un nombre no es solo un rango.** El mismo elemento guarda constantes, fórmulas, referencias a
   varias áreas, referencias externas y nombres integrados como `_xlnm.Print_Area`. El catálogo los
   lista, los marca como «no es un rango» y los conserva al guardar, en vez de descartarlos.
3. **Excel guarda el destino en forma absoluta.** La API recibe `A1:B2` y el archivo guarda
   `$A$1:$B$2`, con la hoja delante cuando el ámbito es el libro y entrecomillada cuando el nombre
   de la hoja lo necesita (`'Mi Hoja'!$A$1`).
4. **La validación de caracteres XML era código muerto.** El nombre solo admite letras, dígitos,
   espacios, `_` y `.`, que son siempre caracteres XML válidos; se eliminó la comprobación
   redundante en vez de dejar una rama imposible.

### Qué **no** se ha implementado en 3.6a (y por tanto no se afirma)

- Referencias 3D, referencias a otros libros, nombres dinámicos y fórmulas con nombre.
- Nombres ocultos, comentarios o ayuda: se conservan, pero no se exponen como propiedades.
- Reordenar hojas, que es lo único que volvería frágil el ámbito posicional.

## Incremento 3.6b — Validación de datos

### Alcance del incremento 3.6b

Escribir y leer las reglas de validación de datos de una hoja. Incluye:

- `ExcelDataValidation.List`, `WholeNumber` y `DecimalNumber` como únicas formas de crear una regla.
- `Worksheet.AddDataValidation`, `Worksheet.DataValidations` y `Worksheet.RemoveDataValidation`.
- `ExcelDataValidationType` con los tipos de SpreadsheetML.
- El operador se deduce de los límites (`between`, `greaterThanOrEqual`, `lessThanOrEqual`).
- Las reglas de otros libros (incluidas las de tipo `custom`) se leen y se conservan.

La decisión completa está en `docs/adr/0027-excel-data-validation.md`.

### Comandos y resultados del 3.6b

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **803 superadas, 0 con error, 0 omitidas** (803 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc36b` | cobertura generada en `artifacts/testresults/phase3/inc36b/**/coverage.cobertura.xml`. |

Las **790 pruebas anteriores siguen pasando**; las 13 nuevas están en `ExcelDataValidationTests`.

### Cobertura del 3.6b (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelDataValidation.cs` | 51 / 52 | 98,1 % | 100 % |
| `Excel/ExcelDataValidations.cs` | 99 / 123 | 80,5 % | 66,4 % |
| **Total de la solución** | 4296 / 4603 | 93,3 % | 84,2 % |

Las líneas sin cubrir son las de lectura de reglas con operadores que esta versión no crea y las
del mapa de tipos que no se pueden alcanzar desde la API pública.

### Verificación independiente del 3.6b (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con reglas de lista y numéricas.
2. **Lectura con el SDK** del tipo, del `sqref`, de las dos fórmulas y del operador
   (`OoxmlValidation.ReadXlsxDataValidations`).
3. **Fixtures construidos con el SDK** (`OoxmlValidation.CreateXlsxWithDataValidations`) con reglas
   de lista, de número entero y personalizada.
4. **Round-trip por stream** del catálogo de reglas.

### Hallazgos del 3.6b durante las pruebas y acciones tomadas

1. **El operador no se puede deducir solo de las fórmulas.** Con un único límite hay que saber si
   es el mínimo o el máximo: si no, la regla se escribía como `between` con una sola fórmula, que
   Excel interpreta de otra forma. El modelo guarda ahora `Minimum` y `Maximum` además de las
   fórmulas, y el operador se deduce de ellos.
2. **Una lista en línea no admite comas.** Excel guarda la lista como una constante entrecomillada
   con los valores separados por comas, así que un valor con coma rompería la regla. La API lo
   rechaza con `ArgumentException` en vez de escribir algo que Excel leería de otra forma.
3. **`<dataValidations>` va después de `<mergeCells>`.** Insertarlo al final rompía el esquema en las
   hojas que ya tenían otros elementos; ahora se coloca por el orden del esquema.
4. **El contenido modelado no se reporta como no modelado.** `dataValidations` dejó de generar el
   diagnóstico `EXCEL-0001` al abrir un libro ajeno.

### Qué **no** se ha implementado en 3.6b (y por tanto no se afirma)

- Reglas de tipo fecha, hora, longitud de texto y fórmula personalizada: se leen y se conservan,
  pero no se pueden crear.
- Los mensajes de error y de entrada: se leen y se conservan, pero las reglas que crea esta versión
  no llevan mensaje.
- Validación de los valores de las celdas: la regla la aplica Excel al escribir, no la librería.

## Incremento 3.6c — Autofiltros

### Alcance del incremento 3.6c

Escribir y leer el autofiltro de una hoja. Incluye:

- `ExcelAutoFilter` con el rango que cubre el filtro.
- `Worksheet.AutoFilter`: lectura, escritura y eliminación (`null`).
- Inserción del elemento en el orden del esquema y fin del diagnóstico de contenido no modelado.

La decisión completa está en `docs/adr/0028-excel-auto-filters.md`.

### Comandos y resultados del 3.6c

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **809 superadas, 0 con error, 0 omitidas** (809 total). |

Las **803 pruebas anteriores siguen pasando**; las 6 nuevas están en `ExcelAutoFilterTests`.

### Cobertura del 3.6c

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelAutoFilter.cs` | 4 / 5 | 80,0 % | 100 % |

La línea sin cubrir es el `ToString`. La cifra de la solución se da en el 3.6d, que es el estado
final de esta tanda: el 3.6c se midió dentro de la misma corrida de cobertura.
### Verificación independiente del 3.6c (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con autofiltro.
2. **Lectura con el SDK** del atributo `ref` del filtro (`OoxmlValidation.ReadXlsxAutoFilter`).
3. **Fixture construido con el SDK** (`OoxmlValidation.CreateXlsxWithAutoFilter`).
4. **Round-trip por stream** del filtro.

### Hallazgos del 3.6c durante las pruebas y acciones tomadas

1. **Los criterios de filtrado son estado, no estructura.** Aplicarlos significa saber qué filas
   están ocultas y con qué condición; eso pertenece a la sesión del usuario, no al archivo. Se
   conservan al guardar la hoja intacta y se descartan al escribir el filtro otra vez.
2. **`<autoFilter>` va antes de `<mergeCells>`.** La inserción respeta el orden del esquema, y una
   hoja con filtro ya no se reporta como contenido no modelado.

### Qué **no** se ha implementado en 3.6c (y por tanto no se afirma)

- Criterios de filtrado, filtros por color y por icono, y estado de ordenación (`sortState`).
- Ocultar filas o columnas: no hay API para ello (queda para el 3.7).

## Incremento 3.6d — Tablas estructuradas

### Alcance del incremento 3.6d

Crear, leer y eliminar tablas estructuradas. Incluye:

- `Worksheet.AddTable(name, rangeReference, hasHeaderRow)`, que crea la parte `xl/tables/tableN.xml`,
  la relación con la hoja y el elemento `<tableParts>`.
- `Worksheet.Tables` y `Worksheet.RemoveTable(name)`.
- `ExcelTable` con `Name`, `Range`, `HasHeaderRow` y `ColumnNames`.
- Nombre de tabla validado y único en el libro; nombres de columna tomados de la cabecera.

La decisión completa está en `docs/adr/0029-excel-structured-tables.md`.

### Comandos y resultados del 3.6d

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **820 superadas, 0 con error, 0 omitidas** (820 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc36d` | **820 superadas**; cobertura generada en `artifacts/testresults/phase3/inc36d/**/coverage.cobertura.xml`. |

Las **809 pruebas anteriores siguen pasando**; las 11 nuevas están en `ExcelTableTests`.

### Cobertura del 3.6d (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelTable.cs` | 10 / 11 | 90,9 % | 100 % |
| `Excel/ExcelTables.cs` | 52 / 56 | 92,9 % | 69,6 % |
| `Excel/ExcelAutoFilter.cs` | 4 / 5 | 80,0 % | 100 % |
| `Excel/Worksheet.cs` | 256 / 270 | 94,8 % | 73,8 % |
| **Total de la solución** | 4462 / 4777 | 93,4 % | 83,7 % |

Las líneas sin cubrir son rutas de error: nombres y rangos que no llegan desde la API pública, y el
`ToString` de los tipos públicos.

### Verificación independiente del 3.6d (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con tabla.
2. **Lectura con el SDK** de la parte de tabla: nombre, rango, `headerRowCount` y columnas
   (`OoxmlValidation.ReadXlsxTables`).
3. **Comprobación del paquete con una herramienta independiente**: la entrada `xl/tables/table1.xml`
   existe y la hoja contiene `<tableParts>` (`TestPackage`).
4. **Fixture construido con el SDK** (`OoxmlValidation.CreateXlsxWithTable`).

### Hallazgos del 3.6d durante las pruebas y acciones tomadas

1. **Una tabla es una parte, no un elemento de la hoja.** Crearla obliga a añadir una parte, su
   relación y el `tablePart` de la hoja; el elemento `<tableParts>` va al final de la hoja, antes de
   `extLst`.
2. **El identificador de la tabla es único en el libro, no en la hoja.** Se calcula como el mayor
   `id` de todas las tablas del libro más uno.
3. **Excel exige nombres de columna únicos.** Dos cabeceras con el mismo texto se desambiguan
   (`Concepto`, `Concepto2`) en lugar de producir una tabla que Excel rechazaría.
4. **Una tabla ajena se conserva y se lee.** Se leen el nombre, el rango, la cabecera y las
   columnas; el resto (estilo, fórmulas de columna) se conserva sin tocar.

### Qué **no** se ha implementado en 3.6d (y por tanto no se afirma)

- Fila de totales, columnas calculadas, fórmulas de columna y referencias estructuradas.
- Elección del estilo de tabla: se usa `TableStyleMedium2` con bandas de fila.
- El autofiltro propio de la tabla no se expone como API.
- No se comprueba que el contenido de la hoja siga siendo coherente con la tabla.

## Incremento 3.7 — Paneles congelados, filas y columnas ocultas, alto y ancho

### Alcance del incremento 3.7

Modelar la vista de la hoja: qué queda congelado al desplazarse, qué filas y columnas están
ocultas y qué tamaño tienen. Incluye:

- `ExcelFreezePanes` y `Worksheet.FreezePanes` (lectura, escritura y eliminación).
- `Worksheet.RowHeight`/`SetRowHeight` e `IsRowHidden`/`SetRowHidden`.
- `Worksheet.ColumnWidth`/`SetColumnWidth` e `IsColumnHidden`/`SetColumnHidden`.
- División de un `<col>` que cubre varias columnas cuando se cambia una de ellas.

La decisión completa está en `docs/adr/0030-excel-sheet-layout.md`.

### Comandos y resultados del 3.7

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **835 superadas, 0 con error, 0 omitidas** (835 total). |

Las **820 pruebas anteriores siguen pasando**; las 15 nuevas están en `ExcelSheetLayoutTests`.

### Cobertura del 3.7

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelFreezePanes.cs` | 20 / 21 | 95,2 % | 100 % |
| `Excel/Worksheet.cs` | 484 / 502 | 96,4 % | 75,3 % |

La línea sin cubrir de `ExcelFreezePanes` es el `ToString`. La cifra de la solución se da en el 3.8,
que es el estado final de estos dos incrementos: el 3.7 se midió dentro de la misma corrida.

### Verificación independiente del 3.7 (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con panel, alto y ancho.
2. **Lectura con el SDK** del panel (`xSplit`, `ySplit`, `topLeftCell`, `state`), de los atributos de
   la fila (`ht`, `hidden`) y de la columna (`width`, `hidden`).
3. **Fixture construido con el SDK** (`OoxmlValidation.CreateXlsxWithSheetLayout`).
4. **Round-trip por stream** del panel, del alto, del ancho y de la visibilidad.

### Hallazgos del 3.7 durante las pruebas y acciones tomadas

1. **Un `<col>` puede cubrir varias columnas.** Cambiar el ancho de la columna 2 de un `<col>` que
   cubre 1..3 no podía tocar ese elemento, porque cambiaría también las columnas vecinas. El
   elemento se divide en tres, y el archivo no queda con columnas solapadas.
2. **Un panel dividido no es un panel congelado.** `xSplit` y `ySplit` significan filas y columnas
   cuando el panel está congelado, pero veinteavos de punto cuando está dividido. `FreezePanes`
   solo informa de los congelados y conserva los otros.
3. **Los elementos modelados dejan de avisar.** `<cols>` y `<sheetViews>` ya no se reportan como
   contenido no modelado.

### Qué **no** se ha implementado en 3.7 (y por tanto no se afirma)

- Paneles divididos, zoom, selecciones visibles y color de pestaña.
- Alto y ancho «automáticos» (`bestFit`) y alto de fila en caracteres.
- Paneles congelados con desfase (`topLeftCell` distinto del que corresponde a los conteos).

## Incremento 3.8 — Configuración de impresión y página

### Alcance del incremento 3.8

Describir cómo se imprime la hoja. Incluye:

- `ExcelPageMargins` y `Worksheet.PageMargins` (márgenes en pulgadas).
- `ExcelPageSetup`, `ExcelPageOrientation`, `ExcelPaperSize` y `Worksheet.PageSetup` (orientación,
  papel, escala y ajuste a páginas).
- `ExcelPrintOptions` y `Worksheet.PrintOptions` (centrado, líneas de la cuadrícula y títulos).
- La marca `pageSetUpPr/@fitToPage` cuando se usa el ajuste a páginas.

La decisión completa está en `docs/adr/0031-excel-print-settings.md`.

### Comandos y resultados del 3.8

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **846 superadas, 0 con error, 0 omitidas** (846 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc38` | **846 superadas**; cobertura generada en `artifacts/testresults/phase3/inc38/**/coverage.cobertura.xml`. |

Las **835 pruebas anteriores siguen pasando**; las 11 nuevas están en `ExcelPrintSettingsTests`.

### Cobertura del 3.8 (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelPageMargins.cs` | 26 / 27 | 96,3 % | 100 % |
| `Excel/ExcelPageSetup.cs` | 29 / 31 | 93,6 % | 95,0 % |
| `Excel/ExcelPrintOptions.cs` | 14 / 15 | 93,3 % | 100 % |
| `Excel/ExcelFreezePanes.cs` | 20 / 21 | 95,2 % | 100 % |
| `Excel/Worksheet.cs` | 484 / 502 | 96,4 % | 75,3 % |
| **Total de la solución** | 4780 / 5104 | 93,7 % | 83,4 % |

Las líneas sin cubrir son los `ToString` de los tipos públicos y el constructor interno que usan las
lecturas.

### Verificación independiente del 3.8 (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con márgenes, configuración y
   opciones de impresión.
2. **Lectura con el SDK** de los márgenes, de la configuración (papel, orientación, escala y ajuste)
   y de las opciones de impresión.
3. **Fixture construido con el SDK** (`OoxmlValidation.CreateXlsxWithPrintSettings`).
4. **Round-trip por stream** de las tres propiedades.

### Hallazgos del 3.8 durante las pruebas y acciones tomadas

1. **La escala y el ajuste a páginas no pueden coexistir.** Escribir las dos produce un archivo que
   Excel resuelve de forma ambigua, así que el constructor lo rechaza.
2. **El ajuste a páginas necesita su marca.** Sin `pageSetUpPr/@fitToPage="1"`, Excel ignora
   `fitToWidth` y `fitToHeight`; la marca se escribe y se conserva al quitarla.
3. **El encabezado y el pie se quedan fuera.** Su formato (`&L`, `&C`, `&R`) es un mini lenguaje
   propio; se conserva en el archivo, pero no hay API.

### Qué **no** se ha implementado en 3.8 (y por tanto no se afirma)

- Encabezado y pie de página.
- Área de impresión, saltos de página, línea de encabezados al repetir y otros atributos de
  `<pageSetup>`.
- Impresión: la configuración describe la hoja, no genera ningún documento.

## Incremento 3.9 — Hipervínculos y notas

### Alcance del incremento 3.9

Escribir y leer los dos elementos de una hoja que no son celdas: los hipervínculos y las notas. Incluye:

- `ExcelHyperlink` y `Worksheet.Hyperlinks`, `AddHyperlink` (externo), `AddInternalHyperlink`
  (interno) y `RemoveHyperlink`, con su relación externa.
- `ExcelComment` y `Worksheet.Comments`, `AddComment` y `RemoveComment`, con la parte de comentarios,
  la lista de autores y el dibujo heredado que Excel necesita.

La decisión completa está en `docs/adr/0032-excel-hyperlinks-and-comments.md`.

### Comandos y resultados del 3.9

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **861 superadas, 0 con error, 0 omitidas** (861 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc39` | **861 superadas**; cobertura generada en `artifacts/testresults/phase3/inc39/**/coverage.cobertura.xml`. |

Las **846 pruebas anteriores siguen pasando**; las 15 nuevas están en `ExcelHyperlinkTests` (7) y
`ExcelCommentTests` (8).

### Cobertura del 3.9 (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelHyperlink.cs` | 18 / 19 | 94,7 % | 0 % |
| `Excel/ExcelComment.cs` | 8 / 9 | 88,9 % | 100 % |
| `Excel/ExcelComments.cs` | 68 / 71 | 95,8 % | 75,0 % |
| **Total de la solución** | 4990 / 5321 | 93,8 % | 83,4 % |

La línea sin cubrir de los tipos públicos es su `ToString`. `ExcelHyperlink` no tiene ramas de
decisión: toda su lógica está en la construcción del objeto.

### Verificación independiente del 3.9 (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con hipervínculo y con notas.
2. **Lectura con el SDK** de los hipervínculos (con la relación resuelta) y de las notas (referencia,
   texto y autor).
3. **Comprobación del paquete con una herramienta independiente**: existen `xl/comments1.xml` y la
   parte de dibujo heredado (`TestPackage`).
4. **Fixtures construidos con el SDK** (`CreateXlsxWithHyperlinks`, `CreateXlsxWithComments`).
5. **Round-trip por stream** de los hipervínculos y de las notas.

### Hallazgos del 3.9 durante las pruebas y acciones tomadas

1. **Un hipervínculo externo es una relación, no un texto.** El destino no va en el elemento
   `<hyperlink>` sino en una relación con `TargetMode="External"`; al eliminar el hipervínculo hay
   que eliminar también la relación.
2. **La dirección se normaliza.** `Uri` añade la barra final a `https://example.com`, así que la
   API devuelve y guarda `https://example.com/`. Se documenta para que el round-trip sea exacto.
3. **Las notas necesitan el dibujo heredado.** Excel no dibuja los indicadores de las notas sin el
   VML, así que se escribe con una forma por nota. La posición del cuadro es aproximada porque
   depende del ancho real de las columnas.
4. **La lista de autores se mantiene al sustituir una nota.** Al volver a escribir la parte de
   comentarios se conservan los autores que ya estaban.

### Qué **no** se ha implementado en 3.9 (y por tanto no se afirma)

- Comentarios con hilo (los de Excel 365) y notas con formato enriquecido.
- Tamaño, visibilidad y color del cuadro de la nota.
- Hipervínculos escritos como fórmula `HYPERLINK()` y vínculos a otros libros.

## Incremento 3.10a — Imágenes en la hoja

### Alcance del incremento 3.10a

Anclar imágenes en una hoja: la parte de medios, el dibujo y la relación. Incluye:

- `ExcelImage`, `ExcelImageFormat` y `Worksheet.AddImage`, `Worksheet.Images` y
  `Worksheet.RemoveImage`.
- Anclaje de una celda (`xdr:oneCellAnchor`) con el tamaño en EMU.

La decisión completa está en `docs/adr/0033-excel-images.md`.

### Comandos y resultados del 3.10a

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **868 superadas, 0 con error, 0 omitidas** (868 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc310a` | **868 superadas**; cobertura generada en `artifacts/testresults/phase3/inc310a/**/coverage.cobertura.xml`. |

Las **861 pruebas anteriores siguen pasando**; las 7 nuevas están en `ExcelImageTests`.

### Cobertura del 3.10a (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelImage.cs` | 17 / 18 | 94,4 % | 100 % |
| `Excel/ExcelImages.cs` | 41 / 43 | 95,4 % | 56,3 % |
| **Total de la solución** | 5108 / 5453 | 93,7 % | 82,7 % |

La línea sin cubrir de `ExcelImage` es su `ToString`. Las ramas de `ExcelImages` sin cubrir son las
del dibujo ajeno que no tiene marcador o extensión.

### Verificación independiente del 3.10a (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con imagen.
2. **Comprobación del paquete con una herramienta independiente**: existen la parte de medios
   (`xl/media/…`) y la de dibujo (`xl/drawings/drawing1.xml`), y la hoja referencia el dibujo.
3. **Fixture construido con el SDK** (`OoxmlValidation.CreateXlsxWithImage`), con el dibujo escrito a
   mano, y lectura de sus propiedades.

### Hallazgos del 3.10a durante las pruebas y acciones tomadas

1. **El relleno de la imagen va en el espacio de nombres de la hoja.** `xdr:blipFill`, no
   `a:blipFill`: con el nombre equivocado el validador del SDK rechazaba el dibujo. Los elementos
   de dentro (`a:blip`, `a:stretch`, `a:fillRect`) sí son del espacio de nombres general.
2. **La imagen no sigue el tamaño de las celdas.** El ancho real de una columna depende de la fuente
   por defecto, así que la imagen se ancla con un tamaño fijo en EMU y no se recalcula.

### Qué **no** se ha implementado en 3.10a (y por tanto no se afirma)

- Anclajes de dos celdas, imágenes flotantes, giros y recortes: se conservan, pero no se crean.
- Validación de que los bytes sean una imagen válida del formato declarado.
- **Gráficos (3.10b)**: quedan fuera de este incremento.

## Incremento 3.10b — Gráficos en la hoja

### Alcance del incremento 3.10b

Escribir gráficos con una serie. Incluye:

- `ExcelChart`, `ExcelChartType` y `Worksheet.AddChart`, `Worksheet.Charts` y
  `Worksheet.RemoveChart`.
- Tipos soportados: columna, barra, línea y sector.
- La parte del gráfico con la referencia a los rangos y su caché de datos.

La decisión completa está en `docs/adr/0034-excel-charts.md`.

### Comandos y resultados del 3.10b

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **885 superadas, 0 con error, 0 omitidas** (885 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase3/inc311` | **885 superadas**; cobertura generada en `artifacts/testresults/phase3/inc311/**/coverage.cobertura.xml`. |

Las **874 pruebas anteriores siguen pasando**; las 11 nuevas están en `ExcelChartTests`.

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Excel/ExcelChart.cs` | 23 / 24 | 95,8 % | 100 % |
| `Excel/ExcelCharts.cs` | 169 / 175 | 96,6 % | 69,1 % |
| `Excel/ExcelImage.cs` | 17 / 18 | 94,4 % | 100 % |
| `Excel/ExcelImages.cs` | 41 / 43 | 95,4 % | 56,3 % |
| **Total de la solución** | 5421 / 5777 | 93,8 % | 82,2 % |

### Verificación independiente del 3.10b (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con gráfico, uno por cada tipo
   soportado.
2. **Comprobación del paquete con una herramienta independiente**: existe la parte del gráfico y la
   del dibujo (`TestPackage`).
3. **Fixture construido con el SDK** (`OoxmlValidation.CreateXlsxWithChart`), con el `chartSpace`
   escrito a mano, que además sirve para comprobar la lectura y la preservación.

### Hallazgos del 3.10b durante las pruebas y acciones tomadas

1. **La parte del gráfico va bajo `xl/drawings/charts/`.** El SDK la coloca junto a la parte que la
   referencia, y no en `xl/charts/` como hace Excel. Es válido: lo que manda es la relación, y el
   validador del SDK lo acepta.
2. **El gráfico necesita dos cosas: la referencia y la caché.** Con la referencia sola, Excel
   recalcularía el gráfico al abrir; con la caché sola, el gráfico no seguiría a los datos. Se
   escriben las dos, y la caché es la foto del momento en que se crea el gráfico.

### Qué **no** se ha implementado en 3.10b (y por tanto no se afirma)

- Varias series, tipos de eje secundario, áreas de combinación, etiquetas de datos y líneas de
  tendencia.
- Gráficos de dispersión, de área, de radar, de cotizaciones o de dispersión: se conservan al
  guardar, pero no se interpretan.
- Recálculo del gráfico al cambiar los datos: la caché no se refresca sola.

## Incremento 3.11 — Tabla de cadenas compartidas

### Alcance del incremento 3.11

Escribir el texto en la tabla de cadenas compartidas, con deduplicación. Incluye:

- `Cell.Value = "texto"` escribe en `xl/sharedStrings.xml` y la celda guarda el índice (`t="s"`).
- El texto repetido se guarda una sola vez.
- Los contadores `count` y `uniqueCount` de la tabla se mantienen al día.
- La lectura de los libros ajenos con tabla de cadenas compartidas se mantiene.

La decisión completa está en `docs/adr/0035-excel-shared-strings.md`.

### Comandos y resultados del 3.11

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **874 superadas, 0 con error, 0 omitidas** (874 total). |

Las **868 pruebas anteriores siguen pasando**; las 6 nuevas están en `ExcelSharedStringTests`. La
prueba de interoperabilidad que comprobaba los *inline strings* ahora comprueba la tabla de cadenas
compartidas.

### Verificación independiente del 3.11 (no autorreferencial)

1. **Validación de esquema** con `OpenXmlValidator` de los paquetes con tabla de cadenas.
2. **Lectura con el SDK** de la tabla: los textos, `count` y `uniqueCount`
   (`OoxmlValidation.ReadXlsxSharedStrings`).
3. **Comprobación del paquete con una herramienta independiente**: existe `xl/sharedStrings.xml` y la
   celda lleva `t="s"`.

### Hallazgos del 3.11 durante las pruebas y acciones tomadas

1. **Cambia el formato de salida de las celdas de texto.** La prueba de interoperabilidad que
   buscaba `inlineStr` falló, y con razón: el incremento cambia precisamente eso.
2. **`count` y `uniqueCount` no son lo mismo.** `uniqueCount` es el número de entradas de la tabla y
   `count` el número de celdas que la referencian. Sustituir una celda de texto por otro valor
   baja el segundo y no el primero, porque la cadena antigua sigue en la tabla (como en Excel).
3. **Los espacios se conservan.** El texto se guarda con `xml:space="preserve"`, para que un texto
   con espacios al final no los pierda al abrirlo.

### Qué **no** se ha implementado en 3.11 (y por tanto no se afirma)

- Reutilización de la tabla al escribir: la tabla no se compacta ni se liberan las cadenas que ya no
  usa ninguna celda.
- Escritura de *inline strings* a petición del llamante: el formato de salida es único.

## Limitaciones conocidas

- El libro completo se mantiene en memoria; el máximo por archivo lo fija
  `OfficeEngineLimits.MaxFileSizeBytes` (256 MiB).
- Toda la I/O es síncrona (ADR 0004).
- Los diagnósticos se calculan al abrir y no se recalculan tras editar.
