# ADR 0025 — Escritura de fórmulas sin cálculo

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.5

## Contexto

En SpreadsheetML una celda con fórmula guarda **dos** cosas: la expresión en `<f>` y el último
resultado conocido en `<v>`. El resultado es una caché: lo escribe el motor de cálculo y solo tiene
sentido mientras las dependencias no hayan cambiado.

Hasta el 3.4 la librería leía el valor almacenado de una celda con fórmula, pero no escribía
fórmulas, y el modelo no distinguía una celda «con fórmula» de una celda «con un valor». Escribir un
valor sobre una celda que ya tenía `<f>` dejaba la fórmula puesta, de modo que el archivo podía
contener una fórmula antigua junto a un valor nuevo.

Escribir fórmulas y calcularlas son dos subsistemas distintos (ADR 0001). Este incremento escribe
fórmulas; no calcula nada. La decisión que hay que tomar es qué hacer con el resultado almacenado:
conservarlo, declararlo ausente o inventarlo.

## Decisión

### 1. `Cell.Formula` es la API de escritura

- El getter devuelve la fórmula con `=` delante (`=A1*B1`) o `null` si la celda no tiene fórmula.
- El setter acepta la fórmula con o sin el `=` inicial y guarda una fórmula normal (`<f>` sin
  atributos de tipo).
- Asignar `null` **elimina** la fórmula y su resultado almacenado de una celda existente; una celda
  ausente sigue ausente. Asignar una fórmula a una celda ausente la crea.
- El formato numérico y el estilo de la celda se conservan.

### 2. Normalización y validación antes de tocar el paquete

- Se elimina un único `=` inicial. El resto del texto se guarda tal cual.
- Se rechaza una fórmula vacía o solo con espacios (`""`, `"="`), una de más de 8192 caracteres y una
  con caracteres que no son válidos en XML.
- La validación ocurre antes de crear la celda: una fórmula inválida deja el libro sin cambios.

### 3. El resultado almacenado se conserva o se declara ausente

- **Al escribir una fórmula** se elimina el resultado almacenado (`<v>`) y el tipo de dato. La
  fórmula queda sin caché: `Cell.Value` devuelve `null` hasta que otro motor la recalcula. No se
  inventa un resultado ni se reutiliza uno que corresponde a otra fórmula.
- **Al leer** una celda con fórmula, `Cell.Value` devuelve el resultado almacenado, que es lo que
  siempre ha hecho. Una fórmula sin resultado almacenado lee `null`.
- **Al abrir y guardar sin tocar la celda**, la fórmula y su resultado se conservan tal cual.
- **Al asignar un valor literal** sobre una celda con fórmula, la fórmula se elimina: una celda no
  puede llevar a la vez un valor asignado por el llamante y una fórmula antigua. `Cell.Value = null`
  elimina ambas cosas.

### 4. Se pide recálculo al abrir

Cuando se escribe o se elimina una fórmula, el libro marca `calcPr/@fullCalcOnLoad="1"`, de forma que
Excel o LibreOffice recalculen al abrir. Si el `calcPr` ya existe, se conservan sus demás atributos
(`calcId`, `calcMode`, iteración…) y solo se cambia esa marca; si no existe, se inserta en el orden
que exige el esquema (`definedNames` antes, `fileRecoveryPr` después).

La librería **no** calcula, no mantiene grafo de dependencias, no publica una matriz de funciones y
no ejecuta la fórmula con `eval` (ADR 0007).

### 5. Lo que no se modela

- Fórmulas **compartidas**, **de matriz** y **de tabla**: se conservan al leer y guardar, pero no se
  escriben ni se expanden. Una fórmula compartida cuyo texto vive en otra celda lee `null`, porque
  esta versión no la expande. Asignar una fórmula sobre una de esas celdas la sustituye por una
  fórmula normal.
- No se valida la semántica de la expresión: `=1+` se escribe tal cual. La responsabilidad de que la
  fórmula sea correcta es del llamante y del motor que la recalcula.
- `CellRange.SetValues` sigue escribiendo valores literales: una cadena que empieza por `=` es
  texto, no una fórmula. `CellRange.GetValues` devuelve el resultado almacenado de las celdas con
  fórmula.

### 6. Errores en la frontera pública

| Situación | Excepción |
| --- | --- |
| Fórmula vacía, de más de 8192 caracteres o con caracteres no válidos en XML | `ArgumentException` |
| Libro o celda almacenada mal formados | `InvalidOfficeFileException` |
| Libro liberado | `ObjectDisposedException` |

## Consecuencias

- La API pública nueva es `Cell.Formula`.
- El paquete producido contiene `<f>` con el texto de la fórmula y, hasta que otro motor recalcula,
  ningún `<v>`.
- Escribir un valor literal ya no deja una fórmula antigua detrás, lo que evita archivos con
  fórmulas y valores contradictorios.
- El libro resultante pide un recálculo completo al abrirse.
- No se afirma ninguna capacidad de cálculo: el motor de fórmulas sigue siendo un subsistema
  separado, tal como se decidió en ADR 0001 y en `docs/roadmap.md`.

## Verificación

- `ExcelFormulaTests`: escritura y round-trip sin resultado almacenado; normalización del `=`
  inicial; ausencia de cálculo; borrado de fórmula y de caché con `null`; limpieza de la caché
  obsoleta al reescribir la fórmula; eliminación de la fórmula al escribir un valor literal y al
  limpiar la celda; validación (vacía, `=`, longitud, carácter no válido en XML) sin crear ni
  modificar la celda; preservación de formato numérico y estilo; petición de recálculo con
  `definedNames` y con un elemento posterior en el orden del esquema; libro liberado.
- `ExcelInteropTests.OpenedWorkbook_WithFormula_ReturnsCachedValueAndPreservesFormula`: una fórmula
  y su resultado de un archivo ajeno se leen y se conservan al guardar.
- `OoxmlValidation.ReadXlsxCellFormula` y `ReadXlsxFullCalculationOnLoad` leen el resultado con el
  SDK, no con la librería bajo prueba; `CreateXlsxWithFormula` construye el fixture con el SDK.
