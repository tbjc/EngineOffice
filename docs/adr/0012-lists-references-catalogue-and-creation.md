# ADR 0012 — Listas: referencias, catálogo y creación

- Estado: Aceptado
- Fecha: 2026-10-17
- Fase: 2, incremento 2.4

## Contexto

El incremento 2.4 implementa «listas numeradas y con viñetas». Se acotó, con aprobación expresa del
interesado, a tres capacidades:

1. **Referencias de lista.** Un párrafo entra en una lista con `w:pPr/w:numPr`, formado por
   `w:ilvl` (nivel, 0..8) y `w:numId` (identificador de la lista). Ambos son `ST_DecimalNumber`
   (enteros). La ausencia de `w:numPr` significa «hereda»; `w:numId="0"` es un valor reservado que
   significa «este párrafo no está numerado», no «sin referencia».
2. **Catálogo de listas de solo lectura.** La parte `word/numbering.xml`
   (`NumberingDefinitionsPart`) contiene `w:numPicBullet*`, `w:abstractNum*` y `w:num*`. Un `w:num`
   referencia una definición abstracta por `w:abstractNumId`; la abstracción `w:abstractNum` guarda
   los `w:lvl`, y cada nivel declara `w:start`, `w:numFmt`, `w:lvlText`, `w:lvlJc` y sus propias
   propiedades de párrafo/carácter.
3. **Creación de listas.** `WordDocument.AddList` crea una lista nueva (un `w:abstractNum` y un
   `w:num`) para que un documento nuevo pueda tener listas funcionales.

Se verificó con una sonda sobre `DocumentFormat.OpenXml` 3.5.1 que:

- El validador **no** comprueba integridad referencial: un `w:numId` que nombra una lista
  inexistente, un documento **sin** `numbering.xml`, `w:numId="0"`, un `w:ilvl="12"` y un `w:num`
  que apunta a un `w:abstractNum` ausente validan con **0 errores**. Word los ignora. Por tanto
  ninguna de esas situaciones es un error.
- Son **errores** de validación: un `w:num` sin `w:numId`, un `w:num` sin `w:abstractNumId` (o con
  el elemento sin `w:val`), un `w:abstractNum` sin `w:abstractNumId` y un `w:lvl` sin `w:ilvl`.
- `CT_Numbering` exige que **todos** los `w:abstractNum` (y `w:numPicBullet`) precedan a **todos**
  los `w:num`.
- `NumberFormatValues` (`ST_NumberFormat`) es una estructura `IEnumValue` con unos 63 miembros. El
  validador rechaza un valor fuera de la enumeración (`w:numFmt="bogus"` → «Enumeration constraint
  failed»), y `LevelJustificationValues` tiene exactamente tres miembros (`left`, `center`, `right`):
  el validador también rechaza `both` y `start`, de modo que un `w:lvlJc` no legible es **inválido**,
  no «válido pero no modelado».
- Los valores mal formados de `w:ilvl`, `w:numId`, `w:numFmt` y `w:lvlJc` lanzan `FormatException` al
  leer la propiedad tipada del SDK.
- El SDK inserta las propiedades tipadas en el orden que exige el esquema, de modo que un `w:lvl`
  construido con `StartNumberingValue`, `NumberingFormat`, `LevelText`, `LevelJustification`,
  `PreviousParagraphProperties` y `NumberingSymbolRunProperties` se serializa en orden válido.
- La propiedad del SDK para `w:num/@w:numId` es `NumberingInstance.NumberID`.

## Decisión

### 1. Referencias tri-state, validadas antes de mutar y nunca resueltas

`Paragraph.ListId` y `Paragraph.ListLevel` son `int?`:

- `null` significa «no se nombra» (elemento ausente): hereda la lista o el nivel.
- `0` en `ListId` es el valor reservado que dice «no numerado»; es distinto de `null` en la API.
- `ListId` rechaza un valor negativo y `ListLevel` un valor fuera de `0..8` con
  `ArgumentOutOfRangeException`, **antes** de tocar el árbol.
- Se escriben en orden de esquema: `w:numPr` en la posición que ya fija `ParagraphPropertyOrder`, y
  dentro de `w:numPr` `w:ilvl` antes de `w:numId` (`NumberingPropertyOrder`). La inserción preserva
  los hijos de `w:numPr` que la versión no modela (`w:numberingChange`, `w:ins`).
- **No** se resuelve la referencia. Office Open XML no tiene integridad referencial y Word ignora
  una referencia colgante; hacerla fallar sería inventar una regla. Se puede resolver con
  `WordDocument.FindList`.
- **No** se toca la definición de la lista.

La lectura usa las propiedades tipadas del SDK; un elemento presente sin `w:val` o con un valor mal
formado lanza `InvalidOfficeFileException`. La lectura informa del valor almacenado aunque quede
fuera del rango que la API escribe (por ejemplo `w:ilvl="12"`), porque el documento es lo que es;
rechazarlo al leer obligaría a inventar.

### 2. Catálogo de solo lectura

`WordDocument.Lists` devuelve las definiciones en orden de documento (lista vacía si no hay parte de
numeración o esta no tiene raíz `w:numbering`) y `WordDocument.FindList(int)` busca por `w:numId`,
devolviendo la primera coincidencia o `null`. `FindList` **omite** una definición cuyo propio
identificador no se pueda leer (una referencia que no se puede leer no puede coincidir).

`WordList` (un `w:num`) expone:

- `Id` (`int`) y `AbstractNumberId` (`int`): **no** nullable, porque el esquema los exige; si faltan
  o no se pueden leer, lanzan `InvalidOfficeFileException`. Esto difiere de `w:styleId`, que el
  validador acepta ausente.
- `Levels` (`IReadOnlyList<WordListLevel>`) y `FindLevel(int)`. Los niveles se leen de la definición
  abstracta referenciada; una referencia colgante deja la lista sin niveles (Word la ignora).

`WordListLevel` (un `w:lvl`) expone `Level` (`int`), `Start` (`int?`), `Format` (`WordListFormat?`),
`Text` (`string?`) y `Justification` (`WordListJustification?`).

### 3. Creación de listas

`WordDocument.AddList(WordListFormat)` crea la parte de numeración si no existe, asigna
identificadores (menor `abstractNumId` libre desde 0; menor `numId` libre desde 1, porque 0 está
reservado), inserta el `w:abstractNum` **antes** del primer `w:num` y añade el `w:num` al final, y
devuelve la `WordList` creada. La lista tiene un solo nivel:

- Nivel 0, `w:start` 1, `w:lvlJc` izquierda, `w:multiLevelType` `singleLevel`.
- `w:pPr/w:ind` con sangría izquierda 720 y francesa 360 twips (media pulgada y cuarto de pulgada),
  como Word.
- Viñeta: `w:numFmt` `bullet`, `w:lvlText` `U+F0B7` y `w:rPr/w:rFonts` `Symbol`, igual que la lista
  con viñetas predeterminada de Word. El resto de formatos: `%1.` como texto.

`Paragraph.ApplyList(WordList, int level = 0)` escribe el identificador y el nivel leyendo antes
`list.Id`, para que una lista ilegible no deje marcado parcial.

### 4. Subconjunto de formatos

`WordListFormat` modela `Bullet`, `Decimal`, `LowerLetter`, `UpperLetter`, `LowerRoman` y
`UpperRoman`, los necesarios para crear una lista corriente. Al **leer** un nivel con un formato
válido pero no modelado (por ejemplo `chicago`) se lanza `UnsupportedOfficeFeatureException`; un
valor mal formado se reporta como `InvalidOfficeFileException`. `WordListJustification` modela
`Left`, `Center` y `Right`, que son todos los valores de `ST_LevelJustification`; un valor que no se
puede leer (por ejemplo `both`, que el validador rechaza) es `InvalidOfficeFileException`.

### 5. Inserción en orden de esquema generalizada

El helper de inserción de `Paragraph` se generaliza a
`InsertInSchemaOrder(OpenXmlCompositeElement, OpenXmlElement, string[])`, con una tabla de orden por
contenedor. `ParagraphPropertyOrder` no cambia; se añade `NumberingPropertyOrder`
(`ilvl`, `numId`, `numberingChange`, `ins`).

### 6. Conservación

Editar el cuerpo no toca `word/numbering.xml` (se conserva byte a byte, comprobado con igualdad
estricta). Añadir una lista solo **añade** elementos: las definiciones existentes quedan intactas,
incluido su marcado no modelado (`w:nsid`, niveles, `w:multiLevelType`…).

### 7. Nomenclatura

La API pública usa el término «lista» (`Lists`, `WordList`, `ListId`, `ListLevel`) y documenta el
mapeo a `w:num`/`w:numPr`/`w:ilvl`; el catálogo nombra los tipos OOXML cuando son relevantes.

## Consecuencias

- Positivas: se pueden crear listas numeradas y con viñetas, situar párrafos en una lista y en un
  nivel, y leer el catálogo (formato, texto y justificación de cada nivel) sin inventar valores; las
  entradas inválidas producen errores tipados y tempranos; el paquete producido es válido; un
  documento sin listas no inventa `numbering.xml`; el `numbering.xml` existente se conserva.
- Limitaciones conocidas:
  - La creación se limita a listas de **un solo nivel**; no se añaden niveles, no se cambia el
    `w:lvlText`, ni `w:startOverride`, ni reinicios, ni `w:numPicBullet`, ni numeración ligada a
    estilos.
  - No se editan ni eliminan definiciones de lista existentes.
  - El modelo de nivel no expone sus propiedades de párrafo/carácter (`w:pPr`/`w:rPr` del nivel);
    solo el formato, el texto, el inicio y la justificación.
  - `WordList`/`WordListLevel` son vistas ligeras ligadas a la vida del `WordDocument`; no se deben
    usar tras liberarlo.
  - El catálogo no genera diagnósticos nuevos: una referencia colgante es válida y no se reporta.
  - La decisión de acotar el incremento a referencias + catálogo + creación (frente a niveles
    personalizados) es del interesado y se refleja en `docs/roadmap.md`.
