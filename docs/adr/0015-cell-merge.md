# ADR 0015 — Tablas: combinación de celdas

- Estado: Aceptado
- Fecha: 2026-10-18
- Fase: 2, incremento 2.5c

## Contexto

El incremento 2.5 «tablas: crear/editar (filas, celdas, bordes, sombreado, ancho, combinación)» se
dividió en tres con aprobación expresa: **2.5a** entregó la estructura y el contenido
(`docs/adr/0013-tables-structure-and-content.md`), **2.5b** las propiedades de tabla y de celda
(`docs/adr/0014-table-and-cell-properties.md`) y este documento fija el **2.5c**: la **combinación
de celdas**.

Hasta ahora, una tabla con celdas combinadas se leía mal y se editaba con cuidado pero sin criterio:
`TableRow.CellCount` informaba del número de celdas **almacenadas**, menor que el número de columnas,
pero no había forma de saber cuántas columnas cubría cada celda ni qué parte jugaba en una
combinación vertical. El marcado (`w:gridSpan`, `w:vMerge`) se **preservaba** sin modelarse.

Se verificó con una sonda sobre `DocumentFormat.OpenXml` 3.5.1 que:

- `w:gridSpan` **exige** `w:val` (ST_DecimalNumber); sin él, el validador del SDK da error de
  esquema. Un `w:val="1"` es válido y equivale a no tener span.
- `w:vMerge` **no** exige `w:val`: ausente significa `continue` (ST_Merge: `restart`/`continue`).
- Una fila con **menos celdas que columnas** declara la rejilla es válida: la rejilla sigue siendo la
  fuente de verdad del número de columnas.
- Una celda de continuación **puede** conservar párrafos: el esquema no lo prohíbe (es una regla de
  semántica de Word, no del esquema).
- Las combinaciones (span solo, span + `vMerge`, rejilla mayor que la suma de spans) **validan** con
  el SDK: el validador no comprueba que los spans sumen las columnas de la rejilla.

## Decisión

### 1. Alcance: dos propiedades y cuatro operaciones

- `TableCell.ColumnSpan` (`int`) — lee y escribe `w:gridSpan`.
- `TableCell.VerticalMerge` (`TableVerticalMerge?`) — lee y escribe `w:vMerge`.
- `TableRow.MergeCells(startIndex, count)` / `TableRow.UnmergeCell(cellIndex)` — combinación
  **horizontal**, que es la que quita o crea celdas.
- `Table.MergeCellsVertically(firstRowIndex, cellIndex, rowCount)` /
  `Table.UnmergeCellsVertically(firstRowIndex, cellIndex, rowCount)` — combinación **vertical**.

Cada operación vive donde vive el dato que toca: la horizontal en la fila (mueve celdas de la fila)
y la vertical en la tabla (recorre filas). `w:hMerge`, la marca **heredada** de combinación
horizontal, no se modela y se preserva.

### 2. `ColumnSpan` es un `int` con 1 por defecto, no un `int?`

Una celda sin `w:gridSpan` cubre una columna y una celda con `w:gridSpan w:val="1"` cubre lo mismo:
son la **misma cosa** expresada de dos formas. Un `int` con valor 1 dice la verdad sin ambigüedad y
no obliga a quien lee a traducir `null`. Coherentemente, asignar `1` **elimina** el elemento (no
escribe un span redundante) y asignar menos de 1 es `ArgumentOutOfRangeException`. Un span
almacenado que falte o no sea un entero ≥ 1 es `InvalidOfficeFileException`: está mal formado.

### 3. `VerticalMerge` es tri-state

`null` significa «no hay `w:vMerge`» (la celda no participa en ninguna combinación), que es
exactamente como OOXML lo expresa. Un `w:vMerge` sin `w:val` lee `Continue`, el valor que el esquema
asume; al escribir, el valor se **declara siempre** (`restart`/`continue`), sin depender del valor
por defecto. Un valor que no sea `continue`/`restart` es `InvalidOfficeFileException`.

### 4. Combinar no destruye contenido

Es la decisión central del incremento y la que más se cuidó.

- **Horizontal** (`MergeCells`): la primera celda de la fila conserva sus propiedades, **gana las
  columnas** de las demás (`span = suma de spans`), y **recibe su contenido** —todos los hijos menos
  `w:tcPr`, en orden, detrás de los suyos—; después las demás celdas se eliminan. Nada se descarta: un
  párrafo, una **tabla anidada** o cualquier otro bloque que un documento ajeno hubiera puesto ahí
  viaja con la celda. Es lo que hace Word al combinar y evita el fallo silencioso más grave posible
  en esta operación.
- **Vertical** (`MergeCellsVertically`): solo escribe las marcas (`restart` en la de arriba,
  `continue` en las de abajo). **No** mueve ni borra contenido. Word concatena el contenido al
  combinar a mano, pero esta versión no lo hace por su cuenta: mover el texto de una celda a otra es
  una decisión del llamante, y hacerlo «por si acaso» sería destruir o duplicar contenido sin que
  nadie lo haya pedido. Queda documentado en el XML doc del método para que nadie se sorprenda.

### 5. El ancho de la celda combinada

La celda que absorbe la combinación declara la **suma** de los anchos de las celdas que cubre cuando
todas declaran un ancho del **mismo tipo** (`dxa` o `pct`); en cualquier otro caso (falta el ancho en
alguna, tipos mezclados, `auto`) la celda combinada se queda **sin ancho preferido** y manda la
rejilla. Se descartó la alternativa barata de dejar el ancho de la primera celda: es una pista
*incorrecta* (media anchura para una celda que ocupa dos columnas) y Word la tiene en cuenta al
recalcular. Leer un ancho que esta versión no modela (la unidad `nil`) durante la combinación falla
con `UnsupportedOfficeFeatureException`, igual que leerlo a mano: la operación no inventa el ancho.

### 6. Guardas contra combinaciones rectangulares y documentos rotos

Las combinaciones rectangulares (una combinación horizontal que a su vez forma parte de una vertical)
no están soportadas. En vez de aceptarlas y producir una tabla que Word no puede reconciliar, las
operaciones **rechazan** el caso con `InvalidOperationException` y no escriben nada:

- `MergeCells`/`UnmergeCell`: si alguna celda del rango tiene `w:vMerge`, la combinación vertical a la
  que pertenece se rompería al cambiar el número de celdas de la fila.
- `MergeCellsVertically`: si la celda de arriba ya es una **continuación** (`continue`), empezar ahí
  una combinación deja huérfana a la de más arriba; y si las celdas del rango cubren distinto número
  de columnas, la combinación no es rectangular.
- `UnmergeCellsVertically` exige el mismo rango mínimo de dos filas que `MergeCellsVertically`: es su
  inversa exacta. Limpiar la marca de **una** celda suelta se hace con
  `TableCell.VerticalMerge = null`, que ya existe.

### 7. Dividir una celda

`UnmergeCell` deja el contenido y las propiedades en la celda original, le quita el span, y la sigue
con tantas celdas vacías como columnas cubría. Las celdas nuevas llevan un `w:p` vacío (el contenido
mínimo que `CT_Tc` espera y el mismo que crea `AddTable`) y **no** declaran ancho: la rejilla vuelve a
mandar. Dividir una celda que ya cubre una columna no crea nada (y limpia un span `1` redundante si lo
hubiera), lo que hace la operación idempotente.

## Consecuencias

- **Nueva capacidad**: `TableCell.ColumnSpan`, `TableCell.VerticalMerge`,
  `TableRow.MergeCells`/`UnmergeCell`, `Table.MergeCellsVertically`/`UnmergeCellsVertically`,
  `TableVerticalMerge`.
- Una tabla combinada ahora **se lee**: `ColumnSpan` dice cuántas columnas cubre cada celda y
  `VerticalMerge` qué parte juega. `TableRow.CellCount` sigue informando de las celdas almacenadas,
  que es un dato distinto y necesario.
- Combinar y dividir **no reescribe** la rejilla (`w:tblGrid`) ni el marcado no modelado
  (`w:hMerge`, `w:noWrap`, `w:tcMar`…); el marcado ajeno del contenedor y de las otras celdas queda
  intacto, y el span se inserta en su sitio de `CT_TcPr` (tras `tcW`, antes de `vMerge`).
- **Fuera de 2.5c**: las combinaciones rectangulares (rechazadas, ver punto 6), los **márgenes de
  celda** (`w:tblCellMar`/`w:tcMar`, aún preservados sin modelar) y `w:hMerge` heredado. Quedan
  anotados en `docs/compatibility-matrix.md`.
- Alternativas descartadas: devolver `int?` en `ColumnSpan` (dos formas de decir «una columna»);
  tomar el contenido de las celdas absorbidas en la combinación vertical (decidir por el llamante
  sobre contenido es destruirlo o duplicarlo); dejar el ancho de la primera celda (pista incorrecta);
  aceptar combinaciones rectangulares y dejar que Word lidie con ellas; mover las operaciones a
  `TableCell` (obligaría a recorrer el árbol para encontrar la fila y la tabla).
