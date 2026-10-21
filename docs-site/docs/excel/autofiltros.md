---
id: autofiltros
title: Autofiltros
sidebar_label: Autofiltros
---

# Autofiltros

Un autofiltro pone los botones de filtro en la cabecera de un rango.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.AutoFilter` | El autofiltro de la hoja, o `null` si no lo tiene. |
| `ExcelAutoFilter.Range` | El rango que cubre el filtro. |

Asignar `null` quita el filtro.

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet["A1"].Value = "Concepto";
sheet["B1"].Value = "Importe";
sheet["A2"].Value = "Ventas";
sheet["B2"].Value = 1250d;

sheet.AutoFilter = new ExcelAutoFilter("A1:B2");

workbook.Save("resumen.xlsx");
```

## Notas

- El rango del filtro incluye la **cabecera**, como en Excel.
- Los criterios que el usuario aplicó (qué filas están ocultas) **no se modelan**: se conservan al
  guardar la hoja intacta, y se descartan al escribir el filtro otra vez.