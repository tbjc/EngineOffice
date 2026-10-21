---
id: parrafos-y-runs
title: Párrafos y runs
sidebar_label: Párrafos y runs
---

# Párrafos y runs

El texto de un párrafo no es una cadena: son **runs**, cada uno con su formato. La biblioteca
distingue las dos cosas.

## API

| Miembro | Qué hace |
| --- | --- |
| `Paragraph.Text` | El texto visible del párrafo. Al asignarlo, el párrafo queda con un único run. |
| `Paragraph.Runs` | Los runs que son hijos directos del párrafo. |
| `Paragraph.AddRun(string text)` | Añade un run al final y lo devuelve. |
| `Run.Text` | El texto del run. |
| `Run.Content` | El contenido real del run: texto, tabulación y saltos. |
| `Run.AddText(string text)` | Añade texto al final del run. |
| `Run.AddTab()` | Añade una tabulación. |
| `Run.AddBreak(BreakType type)` | Añade un salto de línea, de página o de columna. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Paragraph paragraph = document.AddParagraph("Cifras al alza");

Run bold = paragraph.AddRun(" frente al mes anterior.");
bold.Bold = true;

Run underlined = paragraph.AddRun(" Detalle");
underlined.Underline = UnderlineStyle.SingleLine;

document.Save("informe.docx");
```

## Notas

- `Paragraph.Text` **incluye** el texto de los runs que están dentro de un hipervínculo, pero
  `Paragraph.Runs` solo devuelve los que son hijos directos. Para los runs de un hipervínculo está
  `Hyperlink.Runs`.
- Un tabulador o un salto de línea **no se escriben como texto**: se añaden con `Run.AddTab` y
  `Run.AddBreak`. Escribir un `\t` dentro de un texto es un error.