# ADR 0021 — Acceso a celdas por índice de fila y columna

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.1

## Contexto

Hasta este incremento, una celda solo se podía obtener por su referencia A1
(`Worksheet.GetCell(string)` / `worksheet["B7"]`). La conversión interna A1↔índices ya existía
(`CellReference`), pero no se exponía: recorrer un rango exigía construir la referencia como texto en
cada iteración, y una celda no informaba de su propia posición.

La hoja de ruta de la Fase 3 agrupa «Rangos y API por índice de fila/columna» en un solo punto. Son
dos capacidades con decisiones distintas: el acceso por índice no necesita definir qué es un rango ni
cómo se leen y escriben sus valores.

## Decisión

### 1. Índices 1-based, fila primero

`Worksheet.GetCell(int row, int column)` y el indexador `Worksheet[int row, int column]` usan índices
**1-based**, igual que Excel y que las librerías Office: la celda `A1` es `[1, 1]` y `B7` es
`[7, 2]`. El orden de los parámetros es fila y después columna, el mismo que usa Excel al nombrar una
dirección.

El motivo es evitar el error de una unidad que aparece al mezclar la convención del usuario (Excel)
con la del programador (arrays 0-based). La representación almacenada en el paquete sigue siendo la
referencia A1: los índices son una forma de direccionar, no un segundo modelo de datos.

### 2. Límites del formato, validados antes de tocar el paquete

- Fila: `1..1048576`.
- Columna: `1..16384` (columna `XFD`).

Un índice fuera de rango lanza `ArgumentOutOfRangeException` con el nombre del parámetro rechazado
(`row` o `column`). No se recorta ni se crea una celda en otra posición. Los límites salen de
`CellReference.MaxRowIndex`/`MaxColumnIndex`, que ya usaba el analizador de referencias A1.

### 3. Misma semántica que la referencia A1

- El indexador devuelve un **handle vivo**, no un valor.
- Leer un handle cuya celda no existe devuelve `null` y **no crea** la fila ni la celda; asignar un
  valor las crea en orden de columna.
- Un handle obtenido por índice y otro obtenido por A1 para la misma celda describen la misma celda
  del paquete.
- El handle sigue siendo válido mientras el libro no se libere; después, cualquier acceso lanza
  `ObjectDisposedException`.

### 4. La celda informa de su dirección

`Cell.Row` y `Cell.Column` devuelven los índices 1-based de la celda, junto a `Cell.Reference` (A1).
Sirven para recorrer y comparar posiciones sin volver a analizar el texto de la referencia.

### 5. Los rangos quedan para el incremento 3.2

Este ADR no define rangos (`A1:B2`), lectura ni escritura de matrices de valores, ni recorrido de
filas o columnas. Un rango necesita decidir qué devuelve una celda ausente, cómo se tratan los tipos
mezclados y si la escritura crea celdas vacías; son decisiones que no deben colarse en el incremento
del direccionamiento. La hoja de ruta y la matriz reflejan la división.

## Consecuencias

- La API pública nueva es `Worksheet.this[int, int]`, `Worksheet.GetCell(int, int)`, `Cell.Row` y
  `Cell.Column`. La API A1 existente no cambia.
- El paquete producido no cambia: se sigue escribiendo el atributo `r` de `c` con la referencia A1.
- No se añade ninguna dependencia.
- Quedan fuera: rangos, nombres de celda relativos/absolutos (`$A$1`), referencias 3D entre hojas y
  cualquier operación de borrado o movimiento de celdas.

## Verificación

- Pruebas unitarias y de round-trip en `ExcelCellIndexTests` (dirección 1-based, misma celda que A1,
  límites, orden de columnas, celda ausente, libro liberado).
- Validación del paquete con `OpenXmlValidator` y lectura con el SDK de una celda escrita por índice en
  `ExcelInteropTests.ProducedPackage_WithCellWrittenByIndex_IsValidAndReadableByTheSdk`.
