# ADR 0019 — Hipervínculos externos

- Estado: Aceptado
- Fecha: 2026-10-19
- Fase: 2, incremento 2.9

## Contexto

Hasta este incremento, un hipervínculo existente en un `.docx` se preservaba porque el Open XML SDK
conserva el paquete, pero la API no sabía leerlo ni crearlo. Un hipervínculo externo no es solo texto
de un párrafo: es un `w:hyperlink` con un `r:id` que apunta a una **relación de hipervínculo**
(`TargetMode="External"`) de la parte que contiene el párrafo (`word/document.xml`, un encabezado o
un pie).

El criterio del incremento es que el texto y el destino sean legibles y editables, y que sobrevivan
al round-trip sin romper relaciones ajenas. El formato también admite enlaces internos a un marcador
(`w:anchor`) y atributos como `w:tooltip`, `w:history`, `w:tgtFrame` y `w:docLocation`.

## Decisión

### 1. Alcance: solo hipervínculos externos

- `Paragraph.AddHyperlink(string text, Uri destination)` — crea un `w:hyperlink` con una relación
  externa nueva y lo añade al final del párrafo.
- `Paragraph.Hyperlinks` — devuelve los `w:hyperlink` directos cuyo `r:id` resuelve a una relación de
  hipervínculo externa.
- `Hyperlink` — vista viva con `Text`, `Runs` y `Destination`, todos editables.
- Funciona en el cuerpo, en celdas de tabla y en encabezados/pies, porque la parte dueña se descubre
  subiendo al `OpenXmlPartRootElement`, igual que en las imágenes inline (ADR 0018).

Los enlaces internos a marcadores (`w:anchor`) y los enlaces sin destino se preservan pero no se
devuelven. Tampoco se modelan `w:tooltip`, `w:history`, `w:tgtFrame` ni `w:docLocation`: se conservan
tal cual al editar el texto o el destino.

### 2. El destino vive en la relación, no en el texto

`Hyperlink.Destination` lee la `Uri` de la relación. Asignar un valor crea una relación nueva en la
misma parte y actualiza el `r:id` del enlace. No se guarda el destino como texto en el XML, porque
así lo separa Word y porque una relación puede ser relativa o absoluta según el paquete.

La API no valida el esquema de la `Uri` (por ejemplo `http`, `https`, `mailto` o `file`): almacena el
destino y nunca lo abre ni lo ejecuta. Es responsabilidad del llamante decidir qué destinos acepta.

### 3. Relaciones compartidas y relaciones huérfanas

Varios hipervínculos pueden compartir un mismo `r:id`. Si se cambia el destino de uno:

1. se crea una relación nueva y solo ese enlace apunta a ella;
2. la relación anterior se elimina únicamente cuando ya no la usa ningún `w:hyperlink` de la parte.

Así editar un enlace no cambia silenciosamente los demás, y el paquete no acumula relaciones que ya
nadie referencia. El barrido se limita a la parte dueña, porque los identificadores de relación son
locales a la parte.

### 4. Texto y formato

`Hyperlink.Text` concatena todos los `w:t` del enlace, incluidas estructuras anidadas que esta
versión no modela. Asignarlo reemplaza todo el contenido del enlace por un único run y conserva el
destino y los atributos del `w:hyperlink`; el marcado interno que llevaba el texto (marcadores,
formato de carácter, controles) se pierde, igual que al asignar `Paragraph.Text`.

`Hyperlink.Runs` expone los runs directos para editar su texto sin perder el formato. `AddHyperlink`
no aplica el estilo `Hyperlink` ni formato directo: un documento nuevo de OfficeEngine no define ese
estilo y aplicar azul/subrayado a la fuerza sería una decisión visual que la API no debe tomar por el
llamante. Se puede dar formato con `Hyperlink.Runs`.

### 5. Lectura estricta

`Paragraph.Hyperlinks` falla con `InvalidOfficeFileException` si un `w:hyperlink` con `r:id`:

- referencia una relación que no existe;
- referencia una relación que no es de tipo hipervínculo;
- referencia una relación que no es externa.

No se devuelve un enlace a medias ni se inventa un destino. Un `w:hyperlink` sin `r:id` (por ejemplo
un enlace a marcador) no es un error: se preserva y no se devuelve.

## Consecuencias

- API pública nueva: `Paragraph.AddHyperlink`, `Paragraph.Hyperlinks` y `Hyperlink`.
- Añadir un enlace no cambia `Paragraph.Text` salvo por el texto visible del propio enlace; el enlace
  no ocupa un run directo del párrafo, así que `Paragraph.Runs` no lo incluye.
- Editar `Hyperlink.Text` descarta el formato y el contenido interno del enlace, documentado en la
  API.
- No se crean enlaces internos a marcadores, no se elimina ni se mueve un enlace, y no se modelan sus
  atributos adicionales. Todo ello se preserva.
- Cada `AddHyperlink` crea su propia relación; no hay deduplicación por destino.

## Verificación

- Pruebas unitarias de creación, lectura, edición y round-trip por archivo y stream en
  `WordHyperlinkTests`, incluidas celdas de tabla, encabezados/pies, relaciones compartidas, enlaces
  internos preservados y errores tipados.
- Validación con `OpenXmlValidator` y lectura de relaciones y textos con el SDK en
  `WordHyperlinkInteropTests`.
- Fixtures construidos con el SDK (enlace externo, dos enlaces que comparten relación, enlace interno
  y relación ausente o de tipo equivocado) leídos y editados por OfficeEngine.
