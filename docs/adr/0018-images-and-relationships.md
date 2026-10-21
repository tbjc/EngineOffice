# ADR 0018 — Imágenes inline y relaciones

- Estado: Aceptado
- Fecha: 2026-10-19
- Fase: 2, incremento 2.8

## Contexto

Hasta este incremento, una imagen existente en un `.docx` se preservaba porque el SDK conserva el
paquete, pero la API no sabía leerla ni crearla. Una imagen no es solo contenido de un run: es una
**parte** del paquete (`word/media/image1.png`) y una **relación** desde la parte que contiene el
dibujo (`word/document.xml`, un encabezado o un pie) hacia esa parte. En el cuerpo, el dibujo vive
dentro de un `w:drawing`; la variante que este incremento modela es la **imagen inline**
(`wp:inline`) con un `pic:pic` que apunta a la parte mediante `a:blip/@r:embed`.

El criterio del incremento es doble: que una imagen creada o abierta sobreviva al round-trip y que
las relaciones que la versión no modela no se rompan. El SDK ya conserva relaciones y partes si no se
tocan; la responsabilidad de la API es no reemplazar el paquete ni reescribir relaciones ajenas.

## Decisión

### 1. API mínima y explícita

- `Paragraph.AddImage(string path, double widthPoints, double heightPoints)` — infiere el formato de
  la extensión, lee el archivo y añade la imagen al final del párrafo.
- `Paragraph.AddImage(Stream stream, ImageFormat format, double widthPoints, double heightPoints)` —
  para datos que no están en disco; el llamante declara el formato porque el stream no tiene nombre.
- `Paragraph.Images` — devuelve las imágenes inline incrustadas de los runs directos, en orden.
- `Image` — vista viva con `WidthPoints`, `HeightPoints`, `ContentType`, `RelationshipId` y
  `ReadAllBytes()`.
- `ImageFormat` — `Png`, `Jpeg`, `Gif`, `Bmp`, `Tiff`, `Emf` y `Wmf`.

No se añade `Run.AddImage`: una imagen se compone en un **run nuevo** y el punto de entrada natural
es el párrafo. El párrafo conserva su texto y sus runs; la imagen se añade al final.

### 2. Ancho y alto en puntos, convertidos a EMU

El tamaño visible no se deduce de la imagen: el formato no lo exige y decodificar cada formato
obligaría a añadir dependencias o a escribir parsers. El llamante indica el tamaño en puntos y la API
lo convierte a English Metric Units (1 pt = 12 700 EMU), que es la unidad de `wp:extent` y de
`a:xfrm/a:ext`. La validación rechaza valores no finitos, cero, negativos o fuera del rango de
`long` **antes** de crear la parte; una operación inválida no deja una relación huérfana.

### 3. Cada llamada crea su propia parte

No se deduplican bytes. Añadir dos veces la misma imagen crea dos partes y dos relaciones. Es válido
para el formato, evita comparar contenidos potencialmente grandes y mantiene la operación
predecible. La deduplicación queda como una optimización futura, no como una promesa actual.

### 4. La relación se crea en la parte que contiene el dibujo

`Paragraph` descubre su parte dueña subiendo por los ancestros hasta el `OpenXmlPartRootElement`
(`w:document`, `w:hdr` o `w:ftr`). Por eso la misma API funciona en el cuerpo, en una celda de tabla
y en un encabezado o pie, y la relación se crea en la parte correcta (`MainDocumentPart`,
`HeaderPart` o `FooterPart`). El SDK expone `AddImagePart` solo para las partes que declaran esa
relación, así que el tipo se comprueba explícitamente.

### 5. Identificadores de dibujo únicos

`wp:docPr/@id` identifica el objeto de dibujo. Al añadir una imagen se toma el mayor id existente en
el cuerpo, los encabezados y los pies, y se usa el siguiente. Así dos imágenes añadidas en el mismo
párrafo (o en historias distintas) no comparten id.

### 6. Lectura estricta de lo que se modela

`Paragraph.Images` solo devuelve imágenes que puede editar con seguridad. Si un `wp:inline` con un
`pic:pic` incrustado está mal formado, la lectura falla con `InvalidOfficeFileException`: falta
`wp:extent`, el ancho o el alto no son positivos, falta `a:blip`, el `r:embed` está vacío, la
relación no existe o apunta a una parte que no es `ImagePart`. No se inventan dimensiones ni se
devuelve una imagen a medias.

Lo que no se modela **se preserva y no se devuelve**:

- imágenes flotantes (`wp:anchor`);
- imágenes vinculadas (`a:blip/@r:link` y una relación externa);
- gráficos y otros DrawingML que no son `pic:pic`;
- imágenes anidadas en hipervínculos o controles de contenido que la API no modela.

### 7. Relaciones ajenas

La API no reescribe el paquete: al guardar, el SDK conserva las partes y relaciones existentes. La
prueba de interoperabilidad abre un documento con un hipervínculo externo, añade una imagen y
comprueba que la relación del hipervínculo y su `r:id` siguen ahí. También comprueba que una imagen
ajena conserva bytes y relación al editar el cuerpo. Esto satisface el criterio «relaciones no
modeladas no se rompen» sin afirmar que se modelan.

## Consecuencias

- API pública nueva: `Paragraph.AddImage`, `Paragraph.Images`, `Image` e `ImageFormat`.
- Una imagen añadida ocupa un run nuevo; el texto del párrafo no cambia y `Paragraph.Text` sigue
  ignorando el contenido que no es texto.
- Cada imagen añadida crea una parte nueva; no hay deduplicación.
- Fuera de alcance y por tanto no afirmado: imágenes flotantes, imágenes vinculadas, gráficos,
  SVG/JPEG 2000/PCX/iconos al añadir, texto alternativo, tamaño automático a partir de la imagen,
  reemplazo, eliminación o movimiento de imágenes, y validación de que los bytes correspondan al
  formato declarado (no se decodifica la imagen).

## Verificación

- Pruebas unitarias y de round-trip por archivo y stream en `WordImageTests`, incluidas imágenes en
  celdas de tabla y encabezados, validación de argumentos y error tipado ante una relación ausente.
- Validación con `OpenXmlValidator` y lectura de partes, relaciones, bytes y extensiones con el SDK
  en `WordImageInteropTests`.
- Fixtures construidos con el SDK (imagen inline y relación de hipervínculo externa) leídos y
  editados por OfficeEngine, que debe leerlos, preservarlos o fallar con error tipado.
