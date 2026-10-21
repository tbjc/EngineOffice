---
id: hipervinculos
title: Hipervínculos
sidebar_label: Hipervínculos
---

# Hipervínculos

Un hipervínculo externo guarda su **destino** en una relación del paquete, no en el texto.

## API

| Miembro | Qué hace |
| --- | --- |
| `Paragraph.AddHyperlink(string text, Uri destination)` | Añade un hipervínculo y lo devuelve. |
| `Paragraph.Hyperlinks` | Los hipervínculos del párrafo. |
| `Hyperlink.Text` | El texto visible. |
| `Hyperlink.Runs` | Los runs del texto del enlace. |
| `Hyperlink.Destination` | El destino. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Paragraph paragraph = document.AddParagraph("Más información en ");
Hyperlink link = paragraph.AddHyperlink("el sitio de Microsoft", new Uri("https://www.microsoft.com"));
link.Runs[0].Bold = true;

document.Save("informe.docx");
```

## Notas

- Editar el destino crea una **relación nueva** y deja la anterior sin uso: los enlaces que
  compartían esa relación siguen apuntando al sitio viejo.
- Los enlaces **internos** a un marcador (`w:anchor`) se conservan, pero no se crean.
- Los atributos `w:tooltip`, `w:history` y `w:tgtFrame` no se modelan.