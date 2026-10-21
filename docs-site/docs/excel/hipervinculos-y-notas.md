---
id: hipervinculos-y-notas
title: Hipervínculos y notas
sidebar_label: Hipervínculos y notas
---

# Hipervínculos y notas

Un hipervínculo apunta a una dirección (fuera del libro) o a un lugar (dentro). Una nota es un
comentario que se ancla en una celda.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.AddHyperlink(string cellReference, string target, string? display = null, string? tooltip = null)` | Un hipervínculo **externo**. |
| `Worksheet.AddInternalHyperlink(string cellReference, string location, string? display = null, string? tooltip = null)` | Un hipervínculo **interno**. |
| `Worksheet.RemoveHyperlink(string cellReference)` | Elimina el hipervínculo y su relación. |
| `Worksheet.Hyperlinks` | Los hipervínculos de la hoja. |
| `Worksheet.AddComment(string cellReference, string text, string author = "OfficeEngine")` | Añade una nota. |
| `Worksheet.RemoveComment(string cellReference)` | Elimina la nota. |
| `Worksheet.Comments` | Las notas de la hoja. |
| `ExcelHyperlink.Reference`, `Target`, `Location`, `Display`, `Tooltip` | Los datos del hipervínculo. |
| `ExcelComment.Reference`, `Text`, `Author` | Los datos de la nota. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet.AddHyperlink("A1", "https://example.com", display: "Ejemplo", tooltip: "Abre el sitio");
sheet.AddInternalHyperlink("A2", "Resumen!A1", display: "Ir al resumen");
sheet.AddComment("B1", "Revisar este total", author: "Ana");

workbook.Save("resumen.xlsx");
```

## Notas

- Una nota necesita el dibujo heredado que Excel usa para dibujar los indicadores: la biblioteca lo
  escribe, con la posición de la caja **aproximada**.
- La dirección de un hipervínculo externo se guarda en una relación del libro, no en el texto.