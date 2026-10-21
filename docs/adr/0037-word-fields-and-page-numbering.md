# ADR 0037 — Campos y numeración de páginas en Word

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 2, incremento 2.12

## Contexto

Un campo de Word (el número de página, el total de páginas, la fecha, el título del documento) no es
texto: es una **instrucción** que Word evalúa. En el paquete se representa con una secuencia de
runs:

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

El `w:t` del medio es el **resultado almacenado**: lo que ve el lector hasta que Word recalcula el
campo. La librería puede escribir esa secuencia; no puede calcularla.

## Decisión

### 1. API

- `Paragraph.AddField(instruction, cachedResult = null)` — escribe un campo al final del párrafo y
  devuelve un `WordField` con la instrucción y el resultado almacenado.
- `HeaderFooter.AddPageNumber(cachedResult = null)` — escribe un párrafo con el campo `PAGE`.
- `WordField`: expone `Instruction` y `CachedResult`.

### 2. Los campos se escriben, no se evalúan

El resultado almacenado es **el que se da**. La librería no calcula el número de página ni el total
de páginas, no marca el documento para recalcular y no actualiza el resultado al guardar. Word
recalcula el campo al abrir o al imprimir el documento.

El resultado es opcional: sin él, el campo se escribe con los tres runs de apertura, instrucción y
cierre, que es la forma mínima válida.

La instrucción se guarda con los espacios que usa Word (`" PAGE "`), y se valida como texto plano:
no puede contener tabuladores ni saltos de línea.

### 3. El número de página reutiliza el párrafo vacío de la parte

`AddHeader` y `AddFooter` crean la parte con un párrafo vacío (ADR 0017). `AddPageNumber` **reutiliza
ese párrafo** cuando el último de la parte está vacío, para no dejar una línea en blanco encima del
número; si el último párrafo tiene contenido, añade uno nuevo.

### 4. El texto visible no incluye la instrucción

`Paragraph.Text` concatena los `w:t`, y la instrucción de un campo es un `w:instrText`, no un `w:t`:
el texto que muestra el párrafo es el del resultado almacenado. La instrucción solo se ve en el XML.

## Consecuencias

- La API pública nueva es `WordField`, `Paragraph.AddField` y `HeaderFooter.AddPageNumber`.
- El paquete producido contiene campos válidos y el validador del SDK los acepta.
- Un lector que no recalcula campos muestra el resultado almacenado, que puede no ser el número de
  página real.
- Escribir un campo no actualiza ningún resultado: el llamante decide qué resultado guarda.

## Verificación

- `WordFieldTests`: el campo en el pie con su instrucción y sus caracteres de campo; el campo sin
  resultado; el texto alrededor del campo; round-trip por guardado y reopening; que el resultado
  almacenado no se recalcula; instrucciones y resultados inválidos; que el número de página reutiliza
  el párrafo vacío.
- `OoxmlValidation.ReadDocxFooterFields` lee la instrucción y los `fldChar` con el SDK, y
  `ValidateDocx` comprueba el esquema del paquete producido.
