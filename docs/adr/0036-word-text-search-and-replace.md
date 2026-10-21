# ADR 0036 — Búsqueda y reemplazo en Word

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 2, incremento 2.11

## Contexto

En WordprocessingML el texto de un párrafo no es una cadena: son muchos elementos `w:t`, cada uno
dentro de un `w:r` con su propio formato. La palabra «borrador» puede estar repartida en tres runs
—uno en negrita, otro en cursiva y otro con el texto que la precede— y para el usuario es una sola
palabra.

De ahí salen las dos decisiones del incremento:

1. **El alcance de la búsqueda.** Buscar en «el documento» puede significar el cuerpo, o también los
   encabezados, los pies, las notas y los comentarios. Y una búsqueda que cruza el final de un
   párrafo convierte dos párrafos en otro, con efectos que el llamante no pidió.
2. **Qué formato conserva el texto reemplazado.** Si la coincidencia está en varios runs, el texto
   nuevo tiene que quedarse con el formato de alguno de ellos.

## Decisión

### 1. API

- `WordDocument.FindText(text, matchCase = false)` — las coincidencias del **cuerpo**, en orden de
  documento.
- `WordDocument.ReplaceText(oldText, newText, matchCase = false)` — reemplaza y devuelve cuántas
  coincidencias reemplazó.
- `Paragraph.FindText(text, matchCase = false)` y `Paragraph.ReplaceText(oldText, newText,
  matchCase = false)` — el mismo alcance, limitado a un párrafo.
- `WordTextMatch`: expone `Paragraph`, `StartIndex`, `Length` y `Text`.

### 2. El alcance es el cuerpo y no cruza párrafos

- Se buscan los párrafos de primer nivel del cuerpo y los párrafos de las tablas de primer nivel
  (una celda puede tener varios párrafos).
- **No** se buscan los encabezados, los pies, las notas, los comentarios ni los párrafos de tablas
  anidadas: son partes o estructuras que esta versión no modela, y se preservan sin tocar.
- La búsqueda **nunca cruza el final de un párrafo**: «borra» al final de un párrafo y «dor» al
  principio del siguiente no son una coincidencia.

### 3. El texto que se busca

Es la concatenación de los `w:t` de los runs **directos** del párrafo y de los runs de sus
hipervínculos directos. No participan los tabuladores, los saltos de línea, los dibujos ni el texto
de estructuras que no se modelan (controles de contenido, tablas anidadas).

### 4. El texto reemplazado conserva el formato del primer run

Cuando la coincidencia abarca varios runs:

- el texto nuevo se escribe en el **primer** run de la coincidencia, que conserva así su formato de
  carácter;
- el texto de la coincidencia se borra de los runs siguientes, y el último conserva lo que quedaba
  después de la coincidencia;
- el texto anterior y posterior se queda donde estaba.

Es la misma regla que aplica Word, y evita inventar un formato que el documento no tenía.

Las propiedades del párrafo (alineación, sangría, espaciado, estilo, lista) no se tocan, y los runs
que se quedan vacíos se conservan (un `w:r` sin `w:t` es válido).

### 5. Comparación y errores

- La comparación es ordinal y **no distingue mayúsculas** por omisión; con `matchCase: true` distingue.
- Un texto vacío o nulo que se busca lanza `ArgumentException`.
- Un texto de reemplazo nulo lanza `ArgumentNullException`; uno vacío **borra** la coincidencia; uno
  con un tabulador o un salto de línea lanza `ArgumentException`, porque un run no puede storarlos
  (para eso están `Run.AddTab` y `Run.AddBreak`).

Las coincidencias se aplican **de la última a la primera**, para que los desplazamientos de las que
quedan pendientes sigan siendo válidos.

## Consecuencias

- La API pública nueva es `WordTextMatch`, `WordDocument.FindText`, `WordDocument.ReplaceText`,
  `Paragraph.FindText` y `Paragraph.ReplaceText`.
- El alcance es deliberadamente estrecho y está documentado, para que el llamante sepa qué se
  busca y qué no.
- El texto de los encabezados, los pies y las notas no se busca ni se modifica con esta API.

## Verificación

- `WordTextSearchTests`: coincidencias en el cuerpo y su desplazamiento; texto repartido en varios
  runs; comparación con y sin mayúsculas; que la búsqueda no cruce párrafos; que no busque en los
  encabezados; reemplazo de todas las ocurrencias; conservación del formato del primer run;
  conservación del texto alrededor; borrado con texto de reemplazo vacío; reemplazo en tablas;
  round-trip por guardado y reopening; conservación del resto del documento; errores; alcance de
  párrafo.
- Validación con `OpenXmlValidator` y lectura del texto con el SDK en el paquete producido.
