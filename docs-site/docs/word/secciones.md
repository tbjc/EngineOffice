---
id: secciones
title: Secciones
sidebar_label: Secciones
---

# Secciones

Una sección describe el tamaño y la orientación de la página, sus márgenes y el tipo de inicio. Un
documento nuevo tiene una sección US Letter vertical con márgenes de una pulgada.

## API

| Miembro | Qué hace |
| --- | --- |
| `WordDocument.Sections` | Las secciones, en orden de documento. |
| `WordDocument.AddSection(SectionStartType startType)` | Añade una sección al final. |
| `Section.PageSize` | El tamaño y la orientación de la página. |
| `Section.Margins` | Los márgenes. |
| `Section.StartType` | Cómo empieza la sección. |
| `Section.DifferentFirstPage` | Si la primera página tiene su propio encabezado o pie. |

`PageSize` trae los tamaños habituales (`PageSize.Letter`, `PageSize.A4`) y `Rotate()` los gira.
`PageMargins` recibe los cuatro márgenes **en puntos**, en el orden superior, derecho, inferior,
izquierdo.

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Section section = document.Sections[0];
section.PageSize = PageSize.A4;
section.Margins = new PageMargins(72.0, 72.0, 72.0, 72.0);

Section landscape = document.AddSection(SectionStartType.NextPage);
landscape.PageSize = PageSize.A4.Rotate();
landscape.Margins = new PageMargins(36.0, 36.0, 36.0, 36.0);

document.Save("informe.docx");
```

## Notas

- Añadir una sección **copia** el `w:sectPr` anterior, así que la nueva hereda tamaño y márgenes
  hasta que se cambien.
- Las columnas, los bordes de página, la numeración de líneas y la cuadrícula del documento no se
  modelan: se conservan al guardar.