---
id: encabezados-y-pies
title: Encabezados y pies
sidebar_label: Encabezados y pies
---

# Encabezados y pies

Un encabezado o un pie es una **parte** del paquete, no contenido del cuerpo. La sección la referencia
por tipo, y varias secciones pueden compartir la misma parte.

## API

| Miembro | Qué hace |
| --- | --- |
| `Section.Header(HeaderFooterType type)` | El encabezado de ese tipo, o `null` si la sección no lo declara y lo hereda. |
| `Section.Footer(HeaderFooterType type)` | El pie de ese tipo. |
| `Section.AddHeader(HeaderFooterType type)` | Crea el encabezado si no existe y lo devuelve. |
| `Section.AddFooter(HeaderFooterType type)` | Crea el pie si no existe. |
| `Section.RemoveHeader(HeaderFooterType type)` | Quita la referencia; **no borra la parte**. |
| `Section.RemoveFooter(HeaderFooterType type)` | Quita la referencia del pie. |
| `Section.DifferentFirstPage` | `w:titlePg`: la primera página puede tener su propio encabezado o pie. |
| `HeaderFooter.IsHeader`, `Paragraphs`, `Tables`, `Text` | El contenido de la parte. |
| `HeaderFooter.AddParagraph(string text)`, `AddTable(int rows, int columns)` | Añade contenido. |

Los tipos son `Primary` (el de todas las páginas), `FirstPage` y `EvenPage`.

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Section section = document.Sections[0];
section.DifferentFirstPage = true;

HeaderFooter header = section.AddHeader(HeaderFooterType.Primary);
header.Paragraphs[0].Text = "Informe mensual";

HeaderFooter footer = section.AddFooter(HeaderFooterType.Primary);
footer.Paragraphs[0].Text = "Página ";

document.Save("informe.docx");
```

## Notas

- La parte se crea con un párrafo vacío: escríbelo en `Paragraphs[0]` o añade otro con
  `AddParagraph`.
- Quitar la referencia **no borra la parte**: otra sección puede estar usando esa misma parte, y
  borrarla destruiría contenido compartido.
- El tipo `EvenPage` solo se muestra si el documento declara `w:evenAndOddHeaders`, que esta versión
  no modela (se conserva, pero no se escribe).