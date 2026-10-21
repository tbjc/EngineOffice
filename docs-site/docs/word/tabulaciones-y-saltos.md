---
id: tabulaciones-y-saltos
title: Tabulaciones y saltos
sidebar_label: Tabulaciones y saltos
---

# Tabulaciones y saltos

Una tabulación y un salto de línea **no son texto**: tienen su propio elemento en el archivo, y por
eso la biblioteca los ofrece como operaciones.

## API

| Miembro | Qué hace |
| --- | --- |
| `Paragraph.TabStops` | Las tabulaciones del párrafo. |
| `Paragraph.AddTabStop(TabStop stop)` | Añade una tabulación. |
| `Paragraph.ClearTabStops()` | Quita todas las tabulaciones. |
| `Run.AddTab()` | Inserta una tabulación en el run. |
| `Run.AddBreak(BreakType type)` | Inserta un salto de línea, de página o de columna. |
| `Run.Content` | El contenido real del run, en orden. |

`TabStop` se construye con la posición en puntos, la alineación y el relleno:

```csharp
TabStop stop = new TabStop(72.0, TabStopAlignment.Decimal, TabStopLeader.Dot);
```

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Paragraph prices = document.AddParagraph("Precio");
prices.AddTabStop(new TabStop(72.0, TabStopAlignment.Decimal, TabStopLeader.Dot));

// La tabulación y el salto se añaden como contenido, no como texto.
prices.AddRun(string.Empty).AddTab().AddText("1.234,56");

Paragraph page = document.AddParagraph("Continúa en la página siguiente.");
page.AddRun(string.Empty).AddBreak(BreakType.Page);

document.Save("informe.docx");
```

## Notas

- Escribir `"\t"` o `"\n"` dentro de un texto lanza `ArgumentException`: usa `AddTab` y `AddBreak`.
- `BreakType` distingue `TextWrapping` (un salto dentro del párrafo), `Page` y `Column`.