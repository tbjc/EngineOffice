# ADR 0017 — Encabezados y pies por sección

- Estado: Aceptado
- Fecha: 2026-10-19
- Fase: 2, incremento 2.7

## Contexto

El incremento 2.6 modeló la sección (`w:sectPr`), pero dejó sus encabezados y pies preservados sin
modelar. En Office Open XML un encabezado o un pie no es contenido del cuerpo: es una **parte** del
paquete (`word/header1.xml`, `word/footer1.xml`) con raíz `w:hdr`/`w:ftr`, y la sección la referencia
desde `w:headerReference`/`w:footerReference` con un tipo (`w:type`: `default`, `first`, `even`) y una
relación (`r:id`).

Dos hechos del formato mandan sobre la API:

1. **Una parte se puede compartir.** Varias secciones pueden referenciar la misma parte; es lo que
   Word escribe con «vincular con la anterior». Editar la parte desde una sección cambia lo que ven
   las demás.
2. **Una sección sin referencia hereda.** Si una sección no tiene `w:headerReference` de un tipo,
   hereda el de la sección anterior, que es exactamente lo que el formato expresa omitiéndolo.

## Decisión

### 1. Alcance

- `Section.Header(type)` / `Section.Footer(type)` — lee la parte referenciada, o `null` cuando la
  sección no la declara (hereda).
- `Section.AddHeader(type)` / `AddFooter(type)` — crea una parte nueva con un párrafo vacío y su
  referencia, o devuelve la existente. Nunca reemplaza una parte existente.
- `Section.RemoveHeader(type)` / `RemoveFooter(type)` — quita la referencia y **no** borra la parte.
- `HeaderFooter` — vista viva sobre la parte: `IsHeader`, `Paragraphs`, `Tables`, `Text`,
  `AddParagraph` y `AddTable`. Reutiliza los mismos tipos `Paragraph` y `Table` que el cuerpo y la
  misma validación de texto (`PlainText`).
- `Section.DifferentFirstPage` — `w:titlePg`, que decide si Word muestra el encabezado/pie de primera
  página.
- `HeaderFooterType` — `Primary`, `FirstPage`, `EvenPage`, mapeados a `default`, `first`, `even`.

El resto del contenido de un encabezado o pie (controles de contenido y otras estructuras) se
preserva sin modelarse. Las imágenes inline de un encabezado o pie se modelaron en el incremento 2.8
(ADR 0018).

### 2. Quitar la referencia no borra la parte

Podría parecer «más limpio» eliminar la parte al quitar la referencia, pero otra sección puede
seguir apuntando a ella y el formato no exige que toda parte esté referenciada. Borrarla destruiría
contenido compartido. La referencia se quita y la parte queda en el paquete; si nadie la usa, Word la
conserva igualmente. Esto cumple el criterio «vínculos entre sección y parte conservados».

### 3. La creación no rompe el vínculo existente salvo que el llamante lo pida

`AddHeader(type)` solo crea cuando la sección **no** declara ese tipo. Si ya lo declara, devuelve la
parte existente. Así, llamar dos veces no duplica referencias ni pierde contenido, y una sección que
hereda no se materializa hasta que se pide explícitamente.

`WordDocument.AddSection` deja en la sección nueva una copia del `w:sectPr` anterior: las referencias
copian el mismo `r:id`, así que la nueva sección comparte la parte (encabezado o pie) de la anterior,
que es el comportamiento esperado. El ajuste de 2.7 que mueve el elemento original al párrafo de salto
hace además que una vista `Section` obtenida antes de `AddSection` siga describiendo la sección que
terminó, en vez de pasar a describir la nueva.

### 4. Un párrafo vacío al crear

El esquema de `w:hdr`/`w:ftr` exige al menos un elemento de bloque. La parte nueva nace con un
párrafo vacío, que es también lo que Word escribe en un encabezado recién creado. Por eso `Text`
empieza en cadena vacía y `AddParagraph` añade una segunda línea; el llamante puede escribir en
`Paragraphs[0]` directamente.

### 5. Errores tipados, nunca valores inventados

- Referencia sin `w:type` → `InvalidOfficeFileException`.
- `w:type` con un valor que no se puede leer → `InvalidOfficeFileException` (Office Open XML define
  exactamente tres tipos, y los tres se modelan).
- Referencia sin `r:id`, o con una relación que no existe en el paquete, o que apunta a una parte que
  no es un encabezado (`HeaderPart`) o un pie (`FooterPart`) según corresponda → `InvalidOfficeFileException`.
- `w:titlePg` con un valor no booleano → `InvalidOfficeFileException`.
- Un tipo de `HeaderFooterType` no definido → `ArgumentOutOfRangeException` antes de tocar el árbol.

No se validan las referencias al leer otras propiedades de la sección: solo `Header`/`Footer` las
resuelven.

## Consecuencias

- API pública nueva: `HeaderFooter`, `HeaderFooterType`, `Section.Header/Footer/AddHeader/AddFooter/
  RemoveHeader/RemoveFooter` y `Section.DifferentFirstPage`.
- `WordDocument.Text` y `HeaderFooter.Text` comparten el mismo lector de bloques (`BlockText`), así
  que ambos informan el mismo texto para el mismo contenido.
- Quedan fuera: `w:evenAndOddHeaders` (necesario para que Word muestre el tipo `EvenPage`; se
  preserva), encabezados/pies de páginas pares por documento, y el contenido no modelado de las
  partes (se preserva).

## Verificación

- Pruebas unitarias y de round-trip por archivo y stream en `WordHeaderFooterTests`.
- Validación con `OpenXmlValidator` y lectura de referencias y partes con el SDK en
  `WordHeaderFooterInteropTests`.
- Fixtures construidos con el SDK (dos secciones que comparten una parte, parte huérfana, contenido
  no modelado, referencias mal formadas) leídos por OfficeEngine.
