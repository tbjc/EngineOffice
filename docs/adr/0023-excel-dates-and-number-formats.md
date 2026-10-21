# ADR 0023 — Fechas, horas y formatos numéricos

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.3

## Contexto

Hasta el incremento 3.2 una celda solo admitía texto, número, booleano o vacío. En un `.xlsx` una
fecha no es un tipo distinto: es un número serial con un formato numérico que le dice a Excel que lo
muestre como fecha u hora. Sin modelar el formato, leer una fecha como `double` es correcto pero
inútil para el llamante, y escribir un `DateTime` no era posible.

El formato numérico no es solo presentación: decide si un serial representa una fecha. Modelarlo
obliga a decidir varias cosas a la vez: qué API se expone, cómo se convierte el serial en los dos
sistemas de fecha de Excel, cómo se crea un `styles.xml` que un libro nuevo no tiene y cómo se
preserva lo que esta versión no entiende.

## Decisión

### 1. `Cell.NumberFormat` lee y escribe el formato efectivo

- El getter devuelve el código efectivo de la celda, por ejemplo `0.00` o `yyyy-mm-dd`.
- Devuelve `General` cuando la celda no tiene un formato explícito ni hereda uno del `cellXf` 0.
  Una celda ausente también devuelve `General` y no se crea al leer.
- Devuelve `null` para un formato integrado que esta versión no modela. En particular, los formatos
  integrados de fecha/hora dependientes de la configuración regional (ids 27–36 y 50–58) se
  reconocen como fecha/hora, pero no tienen un código canónico que exponer.
- El setter acepta `null` y `General` como «volver a General». Un formato no vacío se valida antes
  de tocar el paquete: máximo 255 caracteres y caracteres válidos en XML.
- Asignar un formato distinto de `General` a una celda ausente la crea para que el formato tenga
  dónde vivir; asignar `General` o `null` a una celda ausente no la crea.
- Cambiar el formato conserva las demás propiedades del estilo de la celda (fuente, relleno, borde y
  alineación): se clona el `cellXf` actual y solo se cambia su referencia de formato numérico.
- Una celda sin atributo `s` hereda el formato del `cellXf` 0. El getter resuelve ese formato; volver
  a `General` añade una referencia explícita a un `cellXf` sin formato si el `cellXf` 0 no era
  `General`.

### 2. Los formatos integrados se exponen con un código canónico

Se modelan los ids integrados 0–22, 37–40 y 45–49 con su código canónico de ECMA-376. Por ejemplo,
el id 14 se expone como `mm-dd-yy` y el id 22 como `m/d/yy h:mm`. Es el código del formato, no el
texto que Excel muestra en una configuración regional concreta.

Los ids integrados que no están en esas tablas (por ejemplo 23–26 y 41–44) no se modelan y se leen
como `null`. Los ids 27–36 y 50–58 se reconocen como fecha/hora para convertir el serial, pero su
código también se lee como `null`. Un archivo con uno de estos formatos sigue siendo legible y el
formato se preserva al editar la celda.

### 3. Escribir una fecha exige un formato explícito

Hay dos formas de escribir una fecha:

- `Cell.SetDateTime(DateTime value, string format)` es la forma principal. Valida que el formato
  contenga al menos un token de fecha u hora, valida el rango de la fecha y escribe el serial y el
  formato en una sola llamada.
- `Cell.Value = DateTime` solo se acepta si la celda ya tiene un formato efectivo de fecha/hora. Si
  no lo tiene, lanza `InvalidOperationException`; la librería no adivina un formato. Esto también
  permite escribir una fecha en una celda que usa un formato integrado de fecha/hora cuyo código no
  se expone.

No se convierte `DateTime.Kind`: se guardan los componentes de fecha y hora tal como están. Un
`DateTime` con `Kind = Utc` no se desplaza a la zona local.

### 4. Seriales de Excel: 1900 y 1904

- En el sistema 1900, la fecha mínima escribible es `1900-01-01` (serial 1). El serial 1 se escribe
  como `1`, el 59 (`1900-02-28`) como `59` y el 61 (`1900-03-01`) como `61`.
- Excel conserva el error histórico del año 1900: el serial 60 representa un `1900-02-29` que no
  existe. Al leer, todo el intervalo `[60, 61)` se devuelve como número, no como `DateTime`, para no
  inventar una fecha. Los seriales menores que 1 (incluido el 0) también se devuelven como número.
- En el sistema 1904, el serial 0 es `1904-01-01`. Se respeta `workbookPr@date1904` del libro
  abierto; esta versión no cambia el sistema de fecha de un libro.
- Una fecha anterior al inicio del sistema del libro lanza `ArgumentOutOfRangeException` antes de
  crear o escribir la celda.
- Al escribir se usa el formato de ida y vuelta (`"R"`) del `double` para no perder precisión en la
  fracción del día. Al leer, un serial fuera del rango representable se devuelve como número.

### 5. Detección de fecha/hora en un código de formato

Un formato se considera de fecha/hora si contiene alguno de los tokens `y`, `m`, `d`, `h`, `s`,
`b`, `g` o `e`, o los marcadores `A/P`/`AM/PM`. La detección:

- Ignora el contenido entre comillas y los caracteres escapados con `\`, `_` o `*`.
- Ignora los corchetes de color, condición o configuración regional, salvo los corchetes de tiempo
  transcurrido `[h]`, `[m]` y `[s]`, que sí son fecha/hora.
- No confunde la notación científica `E+`/`E-` con el token de era `e`.
- Trata `General` como no fecha, incluso cuando aparece después de un prefijo de configuración
  regional.

Es una detección heurística del subconjunto que la librería entiende, no un analizador completo del
lenguaje de formatos de Excel. Un formato con marcadores de AM/PM locales que no usen caracteres
ASCII puede no detectarse; en ese caso el valor se lee como número y el formato se preserva.

### 6. Los rangos no escriben fechas sin formato

`CellRange.GetValues` devuelve `DateTime` para las celdas cuyo formato es de fecha/hora.
`CellRange.SetValues` rechaza `DateTime` porque cada celda necesita su propio formato explícito; la
validación previa existente deja el rango sin cambios. Para escribir fechas en un rango se recorre
con `Cells` y se usa `Cell.SetDateTime` en cada celda.

### 7. El estilo se crea solo cuando hace falta y lo existente se preserva

`ExcelWorkbook.Create` sigue sin crear `xl/styles.xml`. El primer formato no-General crea un
stylesheet mínimo válido con una fuente, un relleno, un borde, un `cellStyleXfs`, un `cellXfs` y el
estilo `Normal`. En un libro abierto, el stylesheet existente no se reescribe: solo se añaden
entradas nuevas y se actualizan los contadores.

- Los formatos personalizados usan ids a partir de 164 y se reutilizan cuando el código coincide.
- Un `cellXf` nuevo se reutiliza si ya existe uno con el mismo XML; si no, se añade al final.
- Un stylesheet malformado que impide resolver o añadir el formato falla con
  `InvalidOfficeFileException`; no se descarta contenido en silencio.

## Consecuencias

- La API pública nueva es `Cell.NumberFormat` y `Cell.SetDateTime(DateTime, string)`. `Cell.Value`
  acepta `DateTime` solo con un formato de fecha/hora ya presente.
- El paquete producido puede incluir `xl/styles.xml` y una referencia `s` en la celda; la fecha se
  guarda como serial numérico. Si no quedan ids de formato personalizado, la escritura falla con
  `UnsupportedOfficeFeatureException`.
- `DateTime` deja de ser un tipo no soportado. `CellRange.SetValues` sigue rechazándolo con
  `UnsupportedOfficeFeatureException`.
- No se implementan todavía: un modelo de solo fecha o solo hora, la exposición de códigos integrados
  dependientes de la configuración regional, la validación semántica completa de un código de
  formato, la conversión de `DateTimeKind` y el cambio del sistema de fecha de un libro. Los estilos
  reutilizables se entregan en el 3.4 (ADR 0024) y la escritura de fórmulas en el 3.5 (ADR 0025).
- El formato de una celda con un código no modelado se conserva aunque `NumberFormat` lo devuelva
  como `null`; editar el valor no lo borra.

## Verificación

- `ExcelDateAndNumberFormatTests`: round-trip del serial y del formato; escritura con
  `SetDateTime`; exigencia de formato explícito; sistema 1900 con sus tres fechas de borde; sistema
  1904; serial 0, 0,5, 60 y 60,5 devueltos como número; formatos integrados y personalizados;
  notación científica y literales que no son fecha; preservación de fuente, relleno, borde y
  alineación al cambiar o resetear el formato; celda sin `s` cuyo formato está en el `cellXf` 0 y
  vuelve a `General`; libro liberado.
- `ExcelCellRangeTests`: `GetValues` devuelve `DateTime` y `SetValues` rechaza `DateTime` sin
  escritura parcial.
- `ExcelInteropTests.ProducedPackage_WithDateAndNumberFormats_IsValidAndReadableByTheSdk`: el
  paquete producido pasa `OpenXmlValidator`; el SDK lee el serial `45322`, el número `1234.5` y los
  códigos `yyyy-mm-dd` y `#,##0.00` desde `cellXfs`/`numFmts`.
- `OoxmlValidation.CreateXlsxWithFormattedCell` construye los fixtures con el SDK, no con la
  librería bajo prueba; `ReadXlsxCellStyle` lee el estilo resultante con el SDK.
