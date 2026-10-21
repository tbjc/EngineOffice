# ADR 0010 — Tabulaciones y saltos

- Estado: Aceptado
- Fecha: 2026-10-16
- Fase: 2, incremento 2.2b

## Contexto

Este incremento añade dos capacidades distintas que se apoyan en el mismo vocabulario (`w:tab`,
`w:br`):

1. **Tabulaciones de párrafo.** `w:pPr/w:tabs` contiene uno o más `w:tab`, cada uno con `w:val`
   (alineación, `ST_TabStop`), `w:pos` (posición, `ST_SignedTwipsMeasure`) y `w:leader` opcional
   (`ST_TabTlc`). El esquema define nueve alineaciones y seis rellenos; OfficeEngine modela seis y
   seis respectivamente (`start`, `end` y `num` no se modelan). Una tabulación no es contenido de un
   run: es formato del párrafo, y un tabulador (`w:tab` de un run) salta a la siguiente parada.
2. **Tabuladores y saltos de run.** Dentro de un run, el contenido no es solo `w:t`: `w:tab` es un
   tabulador, `w:br` un salto (`ST_BrType`: página, columna o ajuste de línea, este último también
   por omisión de `w:type`) y `w:cr` un retorno de carro equivalente a un salto de línea. `w:br`
   admite además `w:clear`, que no se modela.

`w:pos` es **con signo**, a diferencia de `w:firstLine`/`w:hanging` (ADR 0009). Un tabulador o un
salto escrito como carácter (`\t`, `\n`, `\r` dentro de `w:t`) es válido para el esquema pero no es
lo que el llamante pide y se renderiza distinto; y algunos caracteres (el nulo, un suplente suelto)
no son representables en XML en absoluto y solo fallan al guardar.

## Decisión

### 1. Tabulaciones como valor, no como propiedades sueltas

La API expone `Paragraph.TabStops` (lista reconstruida en cada acceso, en orden de documento),
`Paragraph.AddTabStop(TabStop)` y `Paragraph.ClearTabStops()`. Una parada es un tipo de valor
`TabStop` (`PositionPoints`, `TabStopAlignment`, `TabStopLeader`), con igualdad y `ToString`, igual
que `RgbColor`. Se eligió un tipo de valor en lugar de tres propiedades paralelas por párrafo
porque las paradas son una colección ordenada y porque un tipo inmutable se puede comparar y probar
sin tocar el paquete. El miembro `TabStopAlignment.Decimal` conserva el nombre del valor OOXML con
una supresión justificada de `CA1720` (precedente: `UnderlineStyle.SingleLine`, ADR 0008).

### 2. Posición en puntos, el signo se respeta

`PositionPoints` se escribe redondeado a veinteavos de punto (twips) como el resto de medidas. A
diferencia de las sangrías (que rechazan el signo negativo como simplificación documentada en el
ADR 0009), una posición de tabulación **acepta valores negativos**: el esquema es explícitamente con
signo y Word los produce, de modo que rechazarlos perdería información. Un valor no finito o que no
quepa en un twips de 32 bits lanza `ArgumentOutOfRangeException` antes de tocar el árbol.

### 3. Inserción en orden de esquema y preservación

`w:tabs` se coloca dentro de `w:pPr` con el `InsertProperty` del ADR 0009, que ya conocía su
posición relativa (entre `w:shd` y `w:spacing`). Añadir una parada reutiliza el `w:tabs` existente y
**añade al final**, de modo que las paradas ajenas y cualquier hijo no modelado de `w:tabs` se
conservan. Las paradas no se reordenan por posición: el esquema no impone orden entre los `w:tab` y
se comprobó con el validador que una parada añadida fuera de orden sigue produciendo un paquete
válido. `ClearTabStops` elimina `w:tabs` entero (el esquema no permite un `w:tabs` vacío) y
poda `w:pPr` si queda sin hijos ni atributos.

### 4. `Run.Content`: vista fiel, `Text` sigue siendo texto plano

`Run.Content` devuelve la secuencia de piezas modeladas en orden de documento: `RunText`, `RunTab`
y `RunBreak`. `Run.Text` y `Paragraph.Text` siguen concatenando solo `w:t`, sin interpretar
tabuladores ni saltos; su documentación lo dice expresamente. El contenido no modelado de un run
(dibujos, campos, `w:noBreakHyphen`, `w:lastRenderedPageBreak`, tabulaciones absolutas…) se omite de
la vista y **se preserva** al guardar.

### 5. Escritura canónica y lectura explícita

- Un salto de ajuste de línea se escribe **omitiendo** `w:type` (es lo que hace Word); un relleno
  `None` se escribe omitiendo `w:leader`.
- `w:cr` (ajeno) se lee como `BreakType.TextWrapping`.
- Una alineación de tabulación válida pero no modelada (`start`, `end`, `num`) lanza
  `UnsupportedOfficeFeatureException`; una parada sin `w:val`/`w:pos` o con un valor mal formado
  lanza `InvalidOfficeFileException`. Nunca se inventa un valor.
- El tipo de un salto se lee con propiedades tipadas del SDK envueltas en `ReadStoredValue`.

### 6. El texto plano no admite tabuladores ni saltos

Todas las entradas de texto (`Run.Text`, `Run.AddText`, `Paragraph.Text`, `Paragraph.AddRun`,
`Paragraph.AppendText`, `WordDocument.AddParagraph`) pasan por `PlainText.EnsureValid`, que rechaza
`\t`, `\n` y `\r` (con `ArgumentException` que remite a `Run.AddTab`/`Run.AddBreak`) y los caracteres
no representables en XML (el nulo y los suplentes sueltos; los pares de suplentes válidos se
aceptan). La validación se ejecuta **antes** de mutar, de modo que una llamada rechazada no deja
marcado parcial. Sin ella, `\u0000` o un suplente suelto no fallaban al asignar sino al guardar, con
una `ArgumentException` cruda de la BCL (comprobado con una sonda).

## Consecuencias

- Positivas: el paquete producido es válido y con el orden de esquema correcto incluso al editar
  párrafos ajenos; no se pueden escribir tabuladores ni saltos «como texto» que luego se renderizan
  distinto; los errores de entrada son tipados y tempranos; el signo negativo de una parada se
  conserva.
- Limitaciones conocidas: no se modelan las paradas heredadas (estilo, numeración, intervalo por
  defecto), `w:clear` de `w:br` ni las alineaciones `start`, `end` y `num` (estas últimas se leen con
  error tipado). `Run.Content` no incluye el contenido no modelado, por lo que dos `RunText`
  consecutivos pueden haber estado separados por marcado que no se devuelve.
