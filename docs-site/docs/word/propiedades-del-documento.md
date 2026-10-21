---
id: propiedades-del-documento
title: Propiedades del documento
sidebar_label: Propiedades del documento
---

# Propiedades del documento

Las propiedades **centrales** del documento (título, autor, fechas…) se leen y se escriben.

## API

| Miembro | Qué hace |
| --- | --- |
| `WordDocument.Properties` | Las propiedades del documento. |
| `Properties.Title`, `Subject`, `Creator`, `Keywords`, `Description` | Los datos de texto. |
| `Properties.LastModifiedBy`, `Revision`, `Category`, `ContentStatus`, `Identifier`, `Language`, `Version` | Los demás datos de texto. |
| `Properties.Created`, `Modified`, `LastPrinted` | Las fechas. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

document.Properties.Title = "Informe mensual";
document.Properties.Creator = "Oficina de administración";
document.Properties.Subject = "Ventas del trimestre";
document.Properties.Keywords = "ventas, informe";
document.Properties.Created = new DateTime(2026, 1, 15, 9, 0, 0, DateTimeKind.Utc);

document.AddParagraph("Informe generado.");
document.Save("informe.docx");
```

## Notas

- Las propiedades se guardan en `docProps/core.xml`, que **se crea solo** al asignar el primer valor.
- Las propiedades **extendidas** (`docProps/app.xml`) y las **personalizadas**
  (`docProps/custom.xml`) no se modelan: se conservan al guardar, pero no se leen ni se escriben.