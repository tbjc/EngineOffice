---
id: formato-de-parrafo
title: Formato de párrafo
sidebar_label: Formato de párrafo
---

# Formato de párrafo

El formato del párrafo (alineación, sangrías, espacio) se aplica al **párrafo entero**, no a sus
runs.

## API

| Miembro | Qué hace |
| --- | --- |
| `Paragraph.Alignment` | La alineación, con `ParagraphAlignment`. |
| `Paragraph.LeftIndentPoints` | La sangría izquierda, en puntos. |
| `Paragraph.RightIndentPoints` | La sangría derecha, en puntos. |
| `Paragraph.FirstLineIndentPoints` | La sangría de la primera línea. |
| `Paragraph.HangingIndentPoints` | La sangría francesa. |
| `Paragraph.SpacingBeforePoints` | El espacio antes del párrafo. |
| `Paragraph.SpacingAfterPoints` | El espacio después del párrafo. |
| `Paragraph.LineSpacingPoints` | El interlineado, en puntos. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Paragraph summary = document.AddParagraph("Resumen del trimestre");
summary.Alignment = ParagraphAlignment.Center;
summary.LeftIndentPoints = 18.0;
summary.SpacingAfterPoints = 6.0;
summary.LineSpacingPoints = 14.0;

Paragraph justified = document.AddParagraph("El resto del informe, justificado.");
justified.Alignment = ParagraphAlignment.Justify;

document.Save("informe.docx");
```

## Notas

- Todas las medidas están en **puntos** (1/72 de pulgada), no en centímetros ni en el formato
  `w:twips` del archivo.
- El interlineado **múltiple** (`lineRule="auto"`) y las sangrías **en caracteres** no se modelan:
  se conservan al guardar, pero no se pueden escribir.