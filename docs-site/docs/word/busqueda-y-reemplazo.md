---
id: busqueda-y-reemplazo
title: Búsqueda y reemplazo
sidebar_label: Búsqueda y reemplazo
---

# Búsqueda y reemplazo

El texto visible de un párrafo está repartido en varios runs, así que «buscar» significa buscar en
el **texto del párrafo**, no en un run.

## API

| Miembro | Qué hace |
| --- | --- |
| `WordDocument.FindText(string text, bool matchCase = false)` | Las coincidencias del **cuerpo**. |
| `WordDocument.ReplaceText(string oldText, string newText, bool matchCase = false)` | Reemplaza y devuelve cuántas sustituciones hizo. |
| `Paragraph.FindText(string text, bool matchCase = false)` | Lo mismo, limitado a un párrafo. |
| `Paragraph.ReplaceText(string oldText, string newText, bool matchCase = false)` | Lo mismo, limitado a un párrafo. |
| `WordTextMatch.Paragraph`, `StartIndex`, `Length`, `Text` | La coincidencia. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Paragraph paragraph = document.AddParagraph("La versión borrador sigue viva");
paragraph.AddRun(" y el texto continúa en otro run.");

int replaced = document.ReplaceText("borrador", "final");
Console.WriteLine($"Sustituciones: {replaced}");

foreach (WordTextMatch match in document.FindText("final"))
{
    Console.WriteLine($"{match.Paragraph} @ {match.StartIndex}");
}

document.Save("informe.docx");
```

## El alcance, exactamente

- Se busca en el **cuerpo**: los párrafos de primer nivel y los párrafos de las tablas de primer
  nivel. **No** se busca en encabezados, pies, notas ni comentarios.
- La búsqueda **no cruza párrafos**: «borra» al final de un párrafo y «dor» al principio del
  siguiente no son una coincidencia.
- El texto buscado es el de los runs **directos** del párrafo y el de sus hipervínculos directos.

## Qué conserva el reemplazo

El texto nuevo se escribe en el **primer run** de la coincidencia, de modo que conserva su formato
—negrita, cursiva, color—; el resto del texto de la coincidencia se borra de los runs siguientes y
lo que queda antes y después se queda donde estaba (ADR 0036).

## Notas

- Un texto vacío o nulo que se busca lanza `ArgumentException`.
- La comparación no distingue mayúsculas salvo que se pida con `matchCase: true`.