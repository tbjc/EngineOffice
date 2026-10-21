# ADR 0008 — Modelo de formato de carácter en los runs

- Estado: Aceptado
- Fecha: 2026-10-16
- Fase: 2, incremento 2.1

## Contexto

`w:rPr` admite decenas de atributos con valores abiertos (`ST_Underline` tiene 18 formas, `w:sz`
acepta cualquier entero de medios puntos, `w:color` mezcla RGB, colores de tema y `auto`). Además
las propiedades tri-state del OOXML distinguen tres estados reales:

- **ausente**: el run hereda de su estilo/ancestro,
- **presente y activa** (`<w:b/>`),
- **presente y desactivada explícitamente** (`<w:b w:val="false"/>`).

Un modelo que solo guarde `bool` no puede representar el tercer caso y escribiría porcentajes
de estilo incorrectos. Por otra parte, el SDK de Office Open XML lanza excepciones de la BCL al
leer valores que no conoce, algo que OfficeEngine no puede dejar pasar por su API pública.

## Decisión

### 1. Propiedades tri-state

`Run.Bold`, `Run.Italic` y `Run.Underline` son anulables (`bool?` / `UnderlineStyle?`).
Asignar `null` **elimina** el elemento de `w:rPr` (heredar); `false`/`None` escribe el valor
desactivado explícitamente. `w:b` y `w:italic` usan `w:val` omitido para «activado», que es lo
que escribe Word.

### 2. Alcance modelado: solo lo que tiene reglas cerradas

Se modelan familia y tamaño de fuente, negrita, cursiva, subrayado y color RGB. **No** se
modelan resaltado, tachado, sombreados ni colores de tema. Lo no modelado **se preserva** tal
cual en el paquete al editar propiedades vecinas, con prueba dedicada
(`WordFormattingInteropTests.OpenedDocument_UnmodelledRunProperties_SurviveEditingModelledOnes`).

### 3. Color RGB explícito, tema/auto como ausencia de color fijo

`Run.Color` devuelve `RgbColor?`. Un color de tema o `auto` **no** tiene un RGB fijo y
depende de la plantilla del lector, así que se devuelve como `null` y el marcado original
(`w:themeColor`, `w:themeShade`, `val="auto"`) se conserva sin cambios. `RgbColor` es una
estructura de solo lectura con precisión exacta de 24 bits y validación en el constructor.

### 4. Precisión de medio punto

El tamaño es `double?` en puntos. OfficeEngine solo escribe múltiplos de `0.5`
(precisión de `ST_HpsMeasure`); un valor fuera de eso o no finito/positivo lanza
`ArgumentOutOfRangeException`. Un tamaño almacenado no numérico lanza
`InvalidOfficeFileException` al leer.

### 5. Orden de esquema al escribir

`CT_RPr` exige un orden concreto de hijos. Se inserta cada propiedad nuevo en su posición
mediante la tabla `RunPropertyOrder`, no con `AppendChild`, para que el paquete producido sea
siempre válido con el SDK (prueba
`WordFormattingInteropTests.ProducedDocument_WritesRunPropertiesInSchemaOrder`).

### 6. Errores tipados en la frontera de lectura

La API pública nunca deja escapar una excepción de la BCL desde un archivo malformado. El SDK
lanza `FormatException` al analizar un `ST_Underline` o un `ST_OnOff` desconocido; eso se
convierte en `InvalidOfficeFileException` mediante `Run.ReadStoredValue`. El comportamiento se
verificó empíricamente con el SDK 3.5.1 para todos los valores ilegibles que OfficeEngine lee:
solo `FormatException`, por lo que solo se captura ese tipo (no se mantiene código muerto).

Además, la comprobación de validez de una asignación se hace **antes** de tocar el árbol, para
que un valor rechazado no deje marcado parcial en el documento (prueba
`WordCharacterFormatTests.Run_Underline_RejectsAValueOutsideTheModelledSetWithoutTouchingTheDocument`).

## Consecuencias

- Positivas: el modelo de tres estados escribe el mismo `w:rPr` que Word; lo no modelado sobrevive
  a la edición; y cualquier archivo ilegible produce un error tipado y explicado.
- Negativas/limitaciones: quien necesite resaltado, tachado o color de tema deberá esperar a un
  incremento posterior; `Run.Color` no puede informar del color de tema que sí verá Word en
  pantalla.
- Nota: `Run.ToOoxmlUnderline` tiene una rama por defecto que el compilador exige pero que es
  inalcanzable desde valores modelados; se prueba asignando un valor entero fuera del enumerado
  por cast, que lanza `ArgumentOutOfRangeException`.
