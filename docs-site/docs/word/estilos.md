---
id: estilos
title: Estilos
sidebar_label: Estilos
---

# Estilos

Un párrafo y un run pueden **nombrar** un estilo (`w:pStyle`, `w:rStyle`). Además, la biblioteca
ofrece un **catálogo de solo lectura** de los estilos que ya tiene el documento.

## API

| Miembro | Qué hace |
| --- | --- |
| `Paragraph.StyleId` | El estilo de párrafo. |
| `Run.StyleId` | El estilo de carácter. |
| `WordDocument.Styles` | Los estilos del documento, en solo lectura. |
| `WordDocument.FindStyle(string styleId)` | Busca un estilo por su identificador. |
| `WordStyle.Id`, `Name`, `Type`, `IsDefault`, `BasedOnId` | Los datos del estilo. |
| `WordStyle.OwnFormatting` | El formato propio del estilo. |
| `WordStyle.EffectiveFormatting` | El formato propio con la herencia `w:basedOn` resuelta. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Paragraph heading = document.AddParagraph("Resumen");
heading.StyleId = "Heading1";

Run emphasis = heading.AddRun(" del trimestre");
emphasis.StyleId = "Emphasis";

document.Save("informe.docx");

// El catálogo es de solo lectura: describe los estilos que ya existen.
using WordDocument reopened = WordDocument.Open("informe.docx");
WordStyle? style = reopened.FindStyle("Heading1");
if (style is not null)
{
    Console.WriteLine($"{style.Name} (base: {style.BasedOnId})");
}
```

## Notas

- La biblioteca **no crea ni edita definiciones de estilo**: la parte `styles.xml` se conserva tal
  cual. Escribir un `StyleId` que no existe es válido (el archivo lo acepta) y no se avisa, porque
  el estilo lo puede definir otra herramienta.
- `EffectiveFormatting` resuelve la herencia `w:basedOn`; `OwnFormatting` no la sigue.