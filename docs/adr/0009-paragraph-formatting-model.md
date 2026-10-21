# ADR 0009 — Modelo de formato de párrafo

- Estado: Aceptado
- Fecha: 2026-10-16
- Fase: 2, incremento 2.2

## Contexto

`w:pPr` (CT_PPr) es un contenedor con tres particularidades que afectan al diseño:

1. **Orden de hijos obligatorio.** El esquema fija la secuencia exacta (`w:pStyle`, `w:keepNext`,
   …, `w:spacing`, `w:ind`, `w:jc`, …, `w:rPr`, `w:sectPr`). Un párrafo abierto de otra herramienta
   puede traer hijos que esta versión no modela; si se añaden propiedades con `AppendChild`, el
   paquete resultante deja de ser válido.
2. **Unidades mezcladas.** Las sangrías (`w:ind`) y los espaciados `w:before`/`w:after` se miden en
   veinteavos de punto (twips). El espaciado de línea es el caso peligroso: `w:line` está en twips
   solo si `w:lineRule` es `exact` o `atLeast`; con `auto` (valor por defecto cuando el atributo no
   está) es un múltiplo de línea donde `240` es un interlineado sencillo.
3. **Valores que la versión no modela.** `ST_Jc` define once valores y OfficeEngine modela cinco;
   `w:ind` admite variantes en caracteres (`w:leftChars`, `w:firstLineChars`) que no se modelan.

Además, la firma de los atributos no es uniforme: `w:firstLine` y `w:hanging` son
`ST_TwipsMeasure` (sin signo) mientras que `w:left` y `w:right` son `ST_SignedTwipsMeasure`. Escribir
un valor negativo en los dos primeros produce un paquete inválido.

## Decisión

### 1. Inserciones en orden de esquema

Las propiedades se insertan con `Paragraph.InsertProperty`, guiado por la tabla `ParagraphPropertyOrder`
que contiene los **nombres locales XML** (`numPr`, no `numberingProperties`) y la secuencia exacta en
la que el SDK genera los hijos de `ParagraphProperties`, verificada con una sonda de reflexión sobre
`DocumentFormat.OpenXml` 3.5.1 (en una primera versión la tabla tenía nombres inexistentes y omitía
`bidi`, `adjustRightInd` y `snapToGrid`, lo que producía paquetes inválidos al editar párrafos
ajenos). El elemento se desengancha antes de insertarlo para que asignar dos veces la misma propiedad
mueva el nodo en vez de duplicarlo. Pruebas:
`WordParagraphFormatInteropTests.ProducedDocument_WritesParagraphPropertiesInSchemaOrder`,
`..._WithNumbering_KeepsItAheadOfTheModelledProperties` y
`..._WithUnmodelledParagraphFlags_KeepsSchemaOrder`.

### 2. Unidades en puntos con redondeo a 0.05

`FirstLineIndentPoints`, `LeftIndentPoints`, `RightIndentPoints`, `HangingIndentPoints`,
`SpacingBeforePoints`, `SpacingAfterPoints` y `LineSpacingPoints` son `double?` en puntos. Al
escribir se redondea a twips (`Math.Round(points * 20, AwayFromZero)`), por lo que la resolución
real es 0.05 puntos. Al leer se convierte de vuelta (`twips / 20.0`).

### 3. Interlineado: no se adivina

`LineSpacingPoints` escribe siempre `w:lineRule="exact"`. En lectura, si el párrafo almacena
`w:lineRule="auto"` (o no almacena la regla, que equivale a `auto`), la propiedad **lanza
`UnsupportedOfficeFeatureException`** en lugar de devolver un número incorrecto. El marcado original
se conserva. Prueba: `WordParagraphFormatTests.Paragraph_StoredMultipleLineSpacing_ThrowsAndIsPreserved`.

### 4. Límites de validación derivados del esquema

Todas las propiedades exigen un valor finito; todas las sangrías y los espaciados `before`/`after`
exigen además `>= 0`, y el interlineado exige `> 0`. **Sangría francesa:** `w:hanging` es
`ST_TwipsMeasure`, así que un valor negativo se rechaza con `ArgumentOutOfRangeException`
(verificado con el validador del SDK 3.5.1: `hanging="-120"` produce un error de esquema). Las
restricciones de signo se comprobaron además empíricamente (`firstLine`, `hanging` sin signo;
`left`, `right` con signo) antes de fijar el contrato.

Además, `PointsToTwips` rechaza cualquier valor que no quepa en un twips de 32 bits
(`|puntos| <= 107 374 182,35`). Un `(int)` sin comprobar sobre un `double` fuera de rango no lanza:
devuelve `-2147483648` (comprobado con un programa mínimo), de modo que sin esta comprobación un
valor absurdo se habría escrito en el paquete como una medida sin sentido. Prueba:
`Paragraph_RejectsInvalidNumericValues`.

### 5. Valores no modelados en lectura

Un valor que el SDK sabe leer pero OfficeEngine no modela (por ejemplo `w:jc w:val="mediumKashida"`)
lanza `UnsupportedOfficeFeatureException`; un valor que el SDK no puede analizar (`w:jc w:val="bogus"`)
lanza `InvalidOfficeFileException`, envolviendo la `FormatException` del SDK con
`Paragraph.ReadStoredValue`. Nunca se devuelve `null` para un valor presente, porque `null` significa
"heredar" y sería una respuesta falsa. Las variantes en caracteres sí se reportan como `null`
(no hay puntos que devolver) y se preservan.

### 6. Limpieza

Asignar `null` elimina el atributo, el elemento contenedor si queda sin atributos, y `w:pPr` si
queda sin hijos ni atributos. Así un párrafo que se formatea y luego se limpia vuelve al paquete
mínimo. Pruebas: `Paragraph_ClearingFormatting_RemovesTheEmptyProperties`,
`Paragraph_ClearingOneIndent_KeepsTheOthers`.

## Consecuencias

- Positivas: el paquete producido es válido incluso cuando se editan párrafos ajenos con contenido
  no modelado; ninguna unidad se interpreta mal en silencio; el contrato de signos coincide con el
  esquema.
- Negativas/limitaciones: quien lea un documento con interlineado múltiple no obtiene un número por
  `LineSpacingPoints` (debe leer el marcado o esperar a que se modele la regla); las variantes en
  caracteres y la sangría negativa de `w:left`/`w:right` no se modelan (se rechaza el signo negativo
  aunque el esquema lo permita, como simplificación documentada).
