# ADR 0022 — Rangos de celdas

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.2

## Contexto

El incremento 3.1 expuso el acceso a una celda por índice 1-based. Seguir recorriendo una zona de la
hoja obligaba a escribir dos bucles y a construir cada dirección por separado, y no había forma de
leer o escribir un bloque de valores de una vez.

Un rango introduce decisiones que el direccionamiento de una celda no necesita: qué sintaxis se
acepta, si las esquinas pueden invertirse, qué ocurre con las celdas ausentes y cómo se escribe una
matriz heterogénea sin dejar el libro a medio modificar.

## Decisión

### 1. Referencia `primera:última`, sin ambigüedades

`Worksheet.GetRange(string)` acepta la forma A1 `B2:D5`:

- Las dos esquinas son obligatorias y se separan con `:`; una sola celda (`A1`) se rechaza porque
  para eso está el indexador de celda.
- Se aceptan minúsculas y se normaliza a mayúsculas (`b2:d5` → `B2:D5`).
- La esquina izquierda debe ser la superior izquierda y la derecha la inferior derecha. Un rango
  invertido (`B2:A1`) se rechaza en vez de reordenarse en silencio.
- No se aceptan referencias absolutas o mixtas (`$A$1`), espacios, nombres de hoja (`Hoja1!A1`),
  columnas o filas completas (`A:A`, `1:1`), uniones ni intersecciones.

`Worksheet.GetRange(int firstRow, int firstColumn, int lastRow, int lastColumn)` ofrece la misma
operación con índices 1-based, para no construir texto cuando los límites se calculan.

### 2. `CellRange` es una vista viva

`CellRange` expone las dos esquinas (`FirstRow`, `FirstColumn`, `LastRow`, `LastColumn`), el número de
filas y columnas y la referencia normalizada. No copia valores ni conserva un estado propio: lee y
escribe el libro cada vez que se le pide.

- `Cells` enumera las celdas en **orden de fila** (de izquierda a derecha y de arriba abajo). Es una
  secuencia perezosa: no materializa la lista y no crea celdas ausentes.
- `GetValues()` devuelve `object?[,]` con `RowCount` filas y `ColumnCount` columnas. Los índices de la
  matriz son **desplazamientos 0-based** desde la esquina superior izquierda, porque un array de .NET
  no puede ser 1-based. Una celda ausente aporta `null`; una fórmula aporta su valor en caché, igual
  que `Cell.Value`.
- `SetValues(object?[,])` exige exactamente esas dimensiones. Valida **todos** los valores antes de
  escribir el primero, de modo que un tipo no soportado deja el rango intacto.

### 3. `null` significa «sin valor», no «crear una celda vacía»

En `SetValues`, un `null` limpia la celda si existe y **no crea** la celda si falta. Para que la
semántica sea una sola, `Cell.Value = null` deja de crear una celda vacía cuando la celda no existe.

Es un cambio de comportamiento respecto a la versión anterior, que creaba la fila y la celda al
asignar `null`. El motivo es la integridad del paquete: un rango de `null` no debe materializar una
matriz de celdas vacías, y «asignar ningún valor» no debería cambiar la estructura del archivo. Leer
una celda ausente ya devolvía `null` sin crearla.

La celda existente que se limpia conserva su elemento y su formato; solo pierde el valor almacenado.

### 4. Errores tipados en la frontera pública

| Situación | Excepción |
| --- | --- |
| Referencia con sintaxis inválida, esquinas invertidas o dimensiones de matriz incorrectas | `ArgumentException` |
| Esquina fuera de `1..1048576` / `1..16384` | `ArgumentOutOfRangeException` con el parámetro rechazado |
| Valor de tipo no soportado en `SetValues` | `UnsupportedOfficeFeatureException` |
| Celda almacenada mal formada al leer | `InvalidOfficeFileException` |
| Libro liberado | `ObjectDisposedException` |

### 5. Fuera de alcance

- Rangos con nombre, referencias 3D entre hojas, uniones, intersecciones y desplazamiento o
  redimensionado del rango.
- Copiar, mover o borrar el contenido de un rango como operación atómica.
- Estilos, formatos numéricos, fechas y fórmulas: se leen y escriben como celdas sueltas.
- Celdas combinadas: un rango las recorre celda a celda y no expande ni contrae la combinación.

### 6. Rendimiento conocido

`GetValues` y `SetValues` recorren el rango celda a celda y cada acceso localiza la fila y la celda
en el `sheetData`. El coste crece de forma superlineal con el tamaño del rango; es aceptable para los
tamaños habituales y no se optimiza sin una medición real (Fase 5). Está anotado en el código como
`shortcut:`.

## Consecuencias

- La API pública nueva es `Worksheet.GetRange(string)`, `Worksheet.GetRange(int, int, int, int)` y
  `CellRange` (`Reference`, esquinas, `RowCount`, `ColumnCount`, `Cells`, `GetValues`, `SetValues`).
- El paquete producido no cambia de formato: las celdas siguen escribiéndose con su referencia A1.
- No se añade ninguna dependencia.
- `Cell.Value = null` ya no crea una celda vacía; la matriz y el informe lo documentan.

## Verificación

- Pruebas unitarias y de round-trip en `ExcelCellRangeTests`: sintaxis, normalización, ambas
  sobrecargas, orden de fila, vista viva, celdas ausentes, matriz de valores, dimensiones, tipos no
  soportados sin escritura parcial, `null`, libro liberado y referencias inválidas. El cambio de
  `Cell.Value = null` se fija en `ExcelWorkbookTests.Cell_AssignNull_DoesNotCreateAMissingCell`.
- Validación del paquete con `OpenXmlValidator` y lectura con el SDK de una matriz escrita con
  `SetValues` en `ExcelInteropTests.ProducedPackage_WithRangeValues_IsValidAndReadableByTheSdk`.
