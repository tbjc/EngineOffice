# ADR 0016 — Secciones, tamaño/orientación y márgenes

- Estado: Aceptado
- Fecha: 2026-10-19
- Fase: 2, incremento 2.6

## Contexto

Hasta este incremento, `w:sectPr` se **preservaba** pero no se modelaba. Un documento creado por
OfficeEngine no declaraba propiedades de sección, y una sección de un documento ajeno con un tamaño de
página, una orientación o unos márgenes concretos no se podía leer ni editar a través de la API.

Office Open XML describe una sección con `w:sectPr`. Las propiedades de la **última** sección son el
`w:sectPr` que cierra `w:body`; las de cada sección anterior están en el `w:pPr/w:sectPr` del párrafo
que la termina. El esquema exige que el `w:sectPr` del cuerpo sea el **último** hijo de `w:body`.

## Decisión

### 1. Alcance: tres propiedades y una operación

- `Section.PageSize` (`PageSize?`) — `w:pgSz`: ancho, alto (`w:w`/`w:h`) y orientación (`w:orient`).
- `Section.Margins` (`PageMargins?`) — `w:pgMar`: márgenes superior, derecho, inferior e izquierdo.
- `Section.StartType` (`SectionStartType?`) — `w:type`: cómo empieza la sección (`nextPage`,
  `continuous`, `evenPage`, `oddPage`, `nextColumn`).
- `WordDocument.Sections` — todas las secciones, en orden de documento.
- `WordDocument.AddSection(startType)` — cierra la última sección con un párrafo que lleva sus
  propiedades y deja un `w:sectPr` final para la nueva, que hereda el tamaño y los márgenes.

Los encabezados/pies se modelaron en el incremento 2.7 (ADR 0017). Las columnas, los bordes de página,
la numeración de líneas y de páginas, la cuadrícula de documento y el resto de hijos de `w:sectPr` se
**preservan** sin modelarse.

### 2. Valores en puntos y grupos atómicos

Los tamaños y márgenes se exponen en puntos, como el resto de medidas de la API, y se escriben en
twips (1/20 de punto): la resolución real es `0,05` puntos. `PageSize` y `PageMargins` son **structs
inmutables** y se asignan como un todo.

El motivo es el esquema: dentro de `w:pgSz` y `w:pgMar` los atributos son obligatorios u opcionales de
forma que no permite escribir «solo el ancho» ni «solo el margen izquierdo» sin inventar el resto. Un
grupo atómico evita dejar markup a medias y hace que asignar `null` signifique exactamente «quitar el
elemento», no «quitar un atributo requerido».

`PageSize.Rotate()` intercambia las dimensiones y alterna la orientación. La orientación se guarda en
`w:orient` y **no** cambia las dimensiones por sí sola; el método existe para que el caso habitual
—pasar de vertical a horizontal— no obligue a recordar el orden de los números.

### 3. Una sección por defecto en los documentos nuevos

`WordDocument.Create` escribe una sección: **US Letter** (12240 × 15840 twips), vertical, con márgenes
de una pulgada (1440 twips) y distancias de encabezado/pie de 720 twips. Así el tamaño de página de un
documento nuevo es determinista en cualquier equipo, y `Sections` nunca está vacío para un documento
creado por la librería.

Un documento abierto que no declare `w:sectPr` sigue siendo válido y `Sections` informa una lista
vacía: no se inventa una sección. `AddSection` crea la sección por defecto en ese caso, sin párrafo de
salto.

### 4. Añadir una sección no pierde contenido

`AddSection` **mueve** el `w:sectPr` que cerraba el cuerpo al párrafo de salto y deja en el cuerpo una
**copia** para la nueva sección. Así la sección que acaba de terminar conserva el mismo elemento, y una
vista `Section` obtenida antes de la llamada sigue describiendo esa misma sección, no la nueva. El
párrafo de salto es un párrafo vacío, que es como Word almacena un salto de sección: aparece en
`Paragraphs` y en `Text` como una línea vacía. La nueva sección empieza con el mismo tamaño y márgenes y
se puede cambiar por separado.

Este orden se ajustó en el incremento 2.7: antes se clonaba el elemento para el párrafo y se reutilizaba
el del cuerpo, de modo que una vista anterior pasaba a describir la sección nueva.

No se implementan borrar, mover ni fusionar secciones: quitar un `w:sectPr` de un párrafo fusiona dos
secciones, y decidir qué propiedades y qué contenido sobreviven es una operación que este incremento no
aborda.

### 5. Orden de esquema y conservación

Las propiedades se insertan siguiendo la secuencia de `CT_SectPr` (`type`, `pgSz`, `pgMar`, …) con el
mismo helper que usan párrafos y tablas. El resto del marcado de la sección queda donde estaba. Al
editar los márgenes se conservan `w:header`, `w:footer`, `w:gutter` y cualquier atributo no modelado.
Al editar el tamaño se conserva `w:code`. El `w:sectPr` final sigue siendo el último hijo del cuerpo
porque `AddParagraph` y `AddTable` insertan **antes** de él.

### 6. Errores tipados, nunca valores inventados

- `w:pgSz` presente sin `w:w` o `w:h`, o con una medida no numérica → `InvalidOfficeFileException`.
- `w:orient` no reconocido → `InvalidOfficeFileException`.
- `w:pgMar` presente sin alguno de los cuatro márgenes, o con una medida no numérica →
  `InvalidOfficeFileException`.
- `w:type` sin `w:val` o con un valor no reconocido → `InvalidOfficeFileException`.
- Un valor definido por OOXML pero no modelado → `UnsupportedOfficeFeatureException` y el marcado
  original se conserva.
- Valores de escritura fuera de rango (dimensiones no positivas, márgenes horizontales negativos,
  desbordamiento de twips, enums no definidos) → `ArgumentOutOfRangeException` **antes** de tocar el
  árbol.

El margen superior e inferior admiten valores negativos porque el esquema los declara con signo; los
horizontales no.

## Consecuencias

- `WordDocument.Create` produce ahora un cuerpo `[w:sectPr]` que crece insertando contenido delante.
  Las pruebas que contaban los hijos del cuerpo se actualizaron para incluir el `w:sectPr` final.
- La API pública nueva es `WordDocument.Sections`, `WordDocument.AddSection`, `Section`,
  `PageSize`, `PageMargins`, `PageOrientation` y `SectionStartType`.
- Los documentos ajenos conservan las secciones que la API no modela; editar una sección no toca las
  demás.
- Quedan fuera: encabezados/pies por sección (2.7), columnas, saltos de sección al borrar o mover
  contenido y la cuadrícula de documento.

## Verificación

- Pruebas unitarias y de round-trip por archivo y stream en `WordSectionTests`.
- Validación del paquete con `OpenXmlValidator` y lectura de los valores con el SDK en
  `WordSectionInteropTests`.
- Fixtures construidos con el SDK (dos secciones, markup no modelado, medidas y orientación mal
  formadas) leídos por OfficeEngine.
