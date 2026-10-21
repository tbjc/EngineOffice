---
id: tablas-estructuradas
title: Tablas estructuradas
sidebar_label: Tablas estructuradas
---

# Tablas estructuradas

Una tabla es un rango con nombre, con cabecera y con columnas con nombre. A diferencia de una tabla de Word,
aquí es una **parte propia** del libro, con su relación.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.AddTable(string name, string rangeReference, bool hasHeaderRow = true)` | Crea la tabla y la registra en la hoja. |
| `Worksheet.Tables` | Las tablas de la hoja. |
| `Worksheet.RemoveTable(string name)` | Elimina la tabla y su parte. |
| `ExcelTable.Name`, `Range`, `HasHeaderRow`, `ColumnNames` | Los datos de la tabla. |

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

ExcelTable table = sheet.AddTable("Ventas", "A1:B2");
Console.WriteLine(string.Join(", ", table.ColumnNames));

workbook.Save("resumen.xlsx");
```

## Notas

- El nombre de la tabla tiene que ser único en el **libro** y no puede parecer una referencia de
  celda.
- Los nombres de las columnas salen de la **cabecera**; las celdas vacías reciben un nombre
  generado (`Column1`, `Column2`…) y las repetidas se desambiguan.
- La fila de totales, las columnas calculadas y el estilo de la tabla no se modelan.