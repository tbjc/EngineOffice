---
id: impresion
title: Impresión
sidebar_label: Impresión
---

# Impresión

La biblioteca describe **cómo se imprimiría** la hoja: márgenes, tamaño y orientación del papel,
escala o ajuste a páginas, y qué se imprime.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.PageMargins` | Los márgenes, en pulgadas. |
| `ExcelPageMargins(left, right, top, bottom, header = 0.3, footer = 0.3)` | Los márgenes. |
| `Worksheet.PageSetup` | La configuración de la página. |
| `ExcelPageSetup(orientation, paperSize, scale, fitToWidth, fitToHeight)` | La configuración. |
| `Worksheet.PrintOptions` | Qué se imprime y cómo se centra. |
| `ExcelPrintOptions(horizontalCentered, verticalCentered, gridLines, headings)` | Las opciones. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet.PageMargins = new ExcelPageMargins(0.7, 0.7, 0.75, 0.75);
sheet.PageSetup = new ExcelPageSetup(
    ExcelPageOrientation.Landscape,
    ExcelPaperSize.A4,
    fitToWidth: 1,
    fitToHeight: 0);
sheet.PrintOptions = new ExcelPrintOptions(
    horizontalCentered: true,
    gridLines: true);

workbook.Save("resumen.xlsx");
```

## Notas

- La **escala** y el **ajuste a páginas** son excluyentes: la escala fija el porcentaje y el ajuste
  deja que el motor decida cuántas páginas hacen falta. Usar las dos a la vez se rechaza.
- Los márgenes y el tamaño del papel están en **pulgadas** y en **puntos**, según el caso.
- Esto describe la impresión; la biblioteca **no imprime**.