# ADR 0011 — Referencias de estilo y catálogo de estilos (solo lectura)

- Estado: Aceptado
- Fecha: 2026-10-17
- Fase: 2, incremento 2.3

## Contexto

El incremento 2.3 se acotó expresamente a dos capacidades, sin crear ni editar definiciones de
estilo:

1. **Referencias de estilo.** `w:pPr/w:pStyle/@w:val` asocia un estilo de párrafo a un párrafo y
   `w:rPr/w:rStyle/@w:val` asocia un estilo de carácter a un run. Ambos son `ST_String` (longitud
   máxima 253). La ausencia del elemento significa «hereda», no «desactivado».
2. **Catálogo de estilos de solo lectura.** La parte `word/styles.xml` contiene cero o más
   `w:style`. Un estilo puede declarar `w:type` (`ST_Style`: `paragraph`, `character`, `table`,
   `numbering`), `w:styleId`, `w:default` (on/off), `w:name`, `w:basedOn/@w:val` y su propio formato
   en `w:pPr`/`w:rPr`. `w:basedOn` forma una cadena de herencia que hay que resolver para saber el
   formato efectivo.

Se verificó con una sonda sobre `DocumentFormat.OpenXml` 3.5.1 que:

- El validador **no** comprueba integridad referencial: un `w:pStyle`/`w:rStyle` que nombra un estilo
   inexistente, o referencias en un documento **sin** `styles.xml`, validan con **0 errores**. Word
   también ignora la referencia colgante. Por tanto una referencia colgante **no** es un error.
- `w:styleId` y `w:type` ausentes se aceptan (0 errores): son opcionales a efectos prácticos.
- `w:basedOn`, `w:pStyle` y `w:rStyle` **presentes** sin `w:val` producen error de validación («The
   required attribute 'val' is missing»).
- Un `w:pStyle`/`w:rStyle` de más de 253 caracteres produce error de validación.
- Un `w:type` con un valor desconocido lanza `FormatException` al leer `.Value` (igual que el resto
  de enumeraciones `IEnumValue` del SDK).
- `StyleRunProperties` y `StyleParagraphProperties` **no** derivan de `RunProperties` ni de
  `ParagraphProperties`; exponen los mismos hijos hoja (`RunFonts`, `Bold`, `Italic`, `Color`,
  `FontSize`, `Underline`; `Justification`, `Indentation`, `SpacingBetweenLines`).
- `StyleValues` (`ST_Style`) es una estructura `IEnumValue`, no un `enum` de C#, con exactamente
  cuatro miembros: `Paragraph`, `Character`, `Table`, `Numbering`.

## Decisión

### 1. Referencias tri-state, validadas antes de mutar y nunca resueltas

`Paragraph.StyleId` y `Run.StyleId` son `string?`: `null` significa heredar (elemento ausente). El
setter:

- Rechaza vacío o solo espacios, más de 253 caracteres y cualquier carácter que no se pueda
  almacenar en un atributo XML (tab, salto de línea, nulo, suplente suelto), con
  `ArgumentException`, **antes** de tocar el árbol (una llamada rechazada no deja marcado parcial).
- Escribe `w:pStyle`/`w:rStyle` en la posición de esquema que ya fijan `ParagraphPropertyOrder` y
  `RunPropertyOrder` (ambos ya incluían `pStyle`/`rStyle`).
- **No** comprueba que el estilo exista. Office Open XML no tiene integridad referencial, el
  validador del SDK acepta una referencia colgante y Word la ignora; hacerla fallar sería inventar
  una regla que el formato no impone. La referencia se guarda tal cual y se puede resolver con
  `WordDocument.FindStyle`.
- **No** modifica `styles.xml`. Igual que el resto de la biblioteca, editar una referencia no toca la
  definición.

La lectura usa la propiedad tipada del SDK; una referencia presente sin `w:val` lanza
`InvalidOfficeFileException`.

### 2. Catálogo de solo lectura

`WordDocument.Styles` devuelve las definiciones en orden de documento (lista vacía si no hay parte de
estilos o esta no tiene raíz `w:styles`) y `WordDocument.FindStyle(string)` busca por `w:styleId` con
comparación ordinal, devolviendo la primera coincidencia o `null`. `WordStyle` expone:

- `Id` (`string?`), `Name` (`string?`), `Type` (`WordStyleType?`), `IsDefault` (`bool`),
  `BasedOnId` (`string?`).
- `OwnFormatting` y `EffectiveFormatting`.

`Id` y `Type` son nullable porque el SDK acepta su ausencia (0 errores). `w:name` o `w:basedOn`
presentes sin `w:val` lanzan `InvalidOfficeFileException`, y un `w:type` mal formado se reporta como
`InvalidOfficeFileException` (envolviendo la `FormatException` del SDK). `WordStyleType` cubre los
cuatro valores de `ST_Style`, así que no hay rama «no soportada»; el `throw` defensivo del mapeo es
hoy inalcanzable y se mantiene por coherencia con `ReadBreakType` (ADR 0010).

### 3. Modelo de formato del estilo

`WordStyleFormatting` es una estructura de solo lectura con doce miembros nullable (tri-state: `null`
= no definido, hereda):

- De carácter: `FontFamily`, `FontSizePoints`, `Bold`, `Italic`, `Color`.
- De párrafo: `Alignment`, `LeftIndentPoints`, `RightIndentPoints`, `FirstLineIndentPoints`,
  `HangingIndentPoints`, `SpacingBeforePoints`, `SpacingAfterPoints`.

Reutiliza `ParagraphAlignment` y `RgbColor` (ADR 0008). El color sigue el ADR 0008: `auto` y los
colores de tema no tienen RGB fijo y se reportan como `null`. **La fusión es por presencia del
elemento**: si el estilo derivado trae `w:color` —aunque sea `auto` o de tema— sustituye el valor
heredado y **no** filtra el RGB del estilo base; para el resto de propiedades, `null` significa «no
definido» y hereda. La familia de fuente solo considera `w:ascii`/`w:hAnsi`: un `w:rFonts` con solo
`w:cs` no cambia la fuente occidental y, por tanto, hereda. El **interlineado queda excluido**, por el
mismo problema de `w:lineRule="auto"` documentado en el ADR 0009.

### 4. Herencia `w:basedOn`

`EffectiveFormatting` construye la cadena desde el ancestro más lejano hasta el propio estilo y aplica
base → derivado, de modo que una propiedad definida por el derivado sobrescribe la heredada. Un
`w:basedOn` que apunta a un estilo inexistente se ignora (como hace Word). Un ciclo —incluida una
referencia a sí mismo— lanza `InvalidOfficeFileException`. `OwnFormatting` informa solo de lo que
define el estilo, sin resolver.

**No** se pliegan `w:docDefaults` ni el estilo por defecto del tipo: la resolución se limita a la
cadena `w:basedOn` (limitación documentada).

### 5. El lector de estilos no comparte código con `Paragraph`/`Run`

Los mapeos puros (alineación OOXML → `ParagraphAlignment`, twips → puntos) se repiten en `WordStyle`
en lugar de extraerse a un helper compartido con `Paragraph`/`Run`, que ya está verificado. Son
funciones sin estado, con mensajes de error propios («un estilo almacena…»), y el coste de mantenerlas
es menor que el riesgo de tocar código en producción ya cubierto al 100 %.

## Consecuencias

- Positivas: se pueden referenciar estilos y leer su formato propio y heredado sin inventar valores;
  las entradas inválidas producen errores tipados y tempranos; el paquete producido es válido; un
  estilo cuyas propiedades no se modelan se sigue viendo en el catálogo y **se preserva** al guardar
  (la parte `styles.xml` se conserva byte a byte).
- Limitaciones conocidas:
  - No se crean, editan ni eliminan definiciones de estilo.
  - No se resuelven `w:docDefaults`, el estilo por defecto del tipo, los estilos vinculados
    (`w:link`), el estilo siguiente (`w:next`) ni las tablas/numeración.
  - El modelo de formato no cubre todas las propiedades de un estilo (resaltado, tachado, bordes,
    `w:tabs`, interlineado…) ni la herencia de propiedades fuera de `w:basedOn`.
  - El catálogo no genera diagnósticos nuevos: una referencia colgante es válida y no se reporta.
  - La decisión de acotar el incremento a referencias + catálogo (frente a crear/editar estilos) es
    del interesado y se refleja en `docs/roadmap.md`.
