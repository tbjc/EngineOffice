# ADR 0033 — Imágenes en la hoja

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.10a

## Contexto

Una imagen en una hoja son tres cosas a la vez:

1. Los **bytes** de la imagen, en una parte propia (`xl/media/image.png`).
2. El **dibujo** (`xl/drawings/drawing1.xml`), que dice dónde está la imagen, cómo se llama y de
   qué tamaño es.
3. La **relación** entre ambos, más la referencia del dibujo en la hoja (`<drawing r:id="..."/>`).

La hoja no puede contener la imagen: solo la apunta. El dibujo usa el vocabulario de DrawingML para
hojas (`xdr`), no el de DrawingML general (`a`): el contenedor del relleno de la imagen es
`xdr:blipFill`, y dentro van los elementos `a:blip`, `a:stretch` y `a:fillRect`.

## Decisión

### 1. API

- `Worksheet.AddImage(cellReference, content, format, widthPoints, heightPoints, name)`: ancla una
  imagen en una celda.
- `Worksheet.Images`: catálogo de solo lectura con las imágenes de la hoja.
- `Worksheet.RemoveImage(name)`: elimina la imagen, su ancla y su parte de medios.
- `ExcelImage`: expone `Name`, `Reference`, `WidthPoints`, `HeightPoints` y `MediaType`.
- `ExcelImageFormat`: PNG, JPEG, GIF, BMP y TIFF.

### 2. Anclaje y tamaño

La imagen se ancla a **una celda** con `xdr:oneCellAnchor`: un marcador de posición (columna y fila,
sin desplazamiento) y una extensión con el tamaño en unidades inglesas métricas (EMU), donde un
punto son 12700 EMU.

El tamaño se fija al insertar la imagen y **no** se recalcula con el tamaño de las celdas: la
biblioteca no conoce el ancho real de las columnas, que depende de la fuente por defecto. La imagen
no crece ni se encoge al cambiar el ancho de la columna.

### 3. Formatos

Se admiten los cinco formatos de la lista. La biblioteca **no** valida que los bytes sean realmente
una imagen de ese formato: copia el contenido tal cual y declara el tipo de medio correspondiente.
Un contenido que no sea una imagen válida produce un archivo que Excel no podrá mostrar, igual que
ocurre con cualquier otro archivo mal formado.

### 4. Lo que no se modela

- Anclajes de dos celdas (`twoCellAnchor`), imágenes flotantes con posición absoluta, giros y
  recortes: se conservan al guardar, pero no se crean ni se modifican.
- Los gráficos (3.10b) y las formas.
- El contenido de la imagen (no se escala ni se vuelve a codificar).

## Consecuencias

- La API pública nueva es `ExcelImage`, `ExcelImageFormat`, `Worksheet.AddImage`,
  `Worksheet.Images` y `Worksheet.RemoveImage`.
- El paquete producido puede contener una parte de medios, una parte de dibujo y su relación con la
  hoja.
- El tamaño de la imagen es el que se pidió, no el que Excel mostraría.

## Verificación

- `ExcelImageTests`: inserción con su parte de medios, su dibujo y su relación, con round-trip; varias
  imágenes; valores inválidos; eliminación; lectura de un archivo ajeno; preservación al guardar;
  libro liberado.
- `OoxmlValidation.CreateXlsxWithImage` construye el fixture con el SDK, con el dibujo escrito a
  mano, y el paquete se comprueba con `TestPackage`.
