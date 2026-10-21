---
id: numeracion-de-paginas
title: Numeración de páginas
sidebar_label: Numeración de páginas
---

# Numeración de páginas

La numeración de páginas es un **campo**: la biblioteca lo escribe, pero **no lo calcula**. Word
recalcula el campo al abrir o al imprimir.

## API

| Miembro | Qué hace |
| --- | --- |
| `HeaderFooter.AddPageNumber(string? cachedResult = null)` | Añade un párrafo con el campo `PAGE`. |
| `Paragraph.AddField(string instruction, string? cachedResult = null)` | Escribe un campo cualquiera (`PAGE`, `NUMPAGES`…). |
| `WordField.Instruction` | La instrucción del campo. |
| `WordField.CachedResult` | El resultado guardado con el campo. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

HeaderFooter footer = document.Sections[0].AddFooter(HeaderFooterType.Primary);
Paragraph page = footer.AddParagraph("Página ");
page.AddField("PAGE", "1");
page.AddRun(" de ");
page.AddField("NUMPAGES", "1");

document.Save("informe.docx");
```

## Notas

- El **resultado guardado** es el que se ve hasta que Word recalcula. Si no se da ninguno, el campo se
  escribe sin resultado.
- `AddField` escribe la secuencia de runs que usa Word: el carácter que abre el campo, la
  instrucción, el que separa el resultado, el resultado y el que lo cierra.
- Un campo **no es texto**: la instrucción no aparece en `Paragraph.Text`, y el resultado sí.
- La biblioteca no evalúa campos, no los actualiza y no los marca para recalcular.