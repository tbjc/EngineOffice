---
id: formato-de-caracter
title: Formato de carácter
sidebar_label: Formato de carácter
---

# Formato de carácter

Cada run lleva su propio formato: familia y tamaño de letra, negrita, cursiva, subrayado y color.
El formato se aplica **al run**, nunca al párrafo.

## API

| Miembro | Qué hace |
| --- | --- |
| `Run.FontFamily` | La familia de la fuente. |
| `Run.FontSizePoints` | El tamaño en puntos. |
| `Run.Bold` | Negrita (`true`, `false` o `null` para «sin definir»). |
| `Run.Italic` | Cursiva. |
| `Run.Underline` | El subrayado, con el enum `UnderlineStyle`. |
| `Run.Color` | El color RGB, con `RgbColor`. |
| `Run.StyleId` | El estilo de carácter que se aplica al run. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Paragraph paragraph = document.AddParagraph(string.Empty);

Run title = paragraph.AddRun("Cifras al alza");
title.FontFamily = "Arial";
title.FontSizePoints = 14.0;
title.Bold = true;
title.Color = new RgbColor(0x1F, 0x49, 0x7D);

Run detail = paragraph.AddRun(" frente al mes anterior.");
detail.Italic = true;
detail.Underline = UnderlineStyle.SingleLine;

document.Save("informe.docx");
```

## Notas

- Las tres propiedades de negrita, cursiva y subrayado son **tri-estado**: `null` significa «sin
  definir» y deja el valor heredado del estilo, mientras que `false` lo desactiva de forma explícita.
- El color se guarda como RGB; el **color de tema** de Word no se modela (se conserva, pero no se
  escribe).