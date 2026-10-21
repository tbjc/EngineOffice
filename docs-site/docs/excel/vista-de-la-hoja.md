---
id: vista-de-la-hoja
title: Vista de la hoja
sidebar_label: Vista de la hoja
---

# Vista de la hoja

Lo que el usuario ve al desplazarse: qué queda congelado, qué filas y columnas están ocultas y
cuánto miden.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.FreezePanes` | Los paneles congelados, o `null` si no los hay. |
| `ExcelFreezePanes(int rows, int columns)` | Crea un panel congelado. |
| `Worksheet.RowHeight(int row)` | El alto de una fila, en puntos. |
| `Worksheet.SetRowHeight(int row, double points)` | Fija el alto de una fila. |
| `Worksheet.IsRowHidden(int row)` | Si la fila está oculta. |
| `Worksheet.SetRowHidden(int row, bool hidden)` | Oculta o muestra una fila. |
| `Worksheet.ColumnWidth(int column)` | El ancho de una columna, en caracteres. |
| `Worksheet.SetColumnWidth(int column, double characters)` | Fija el ancho de una columna. |
| `Worksheet.IsColumnHidden(int column)` | Si la columna está oculta. |
| `Worksheet.SetColumnHidden(int column, bool hidden)` | Oculta o muestra una columna. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

// La primera fila y la primera columna quedan fijas al desplazarse.
sheet.FreezePanes = new ExcelFreezePanes(rows: 1, columns: 1);

sheet.SetColumnWidth(1, 24);
sheet.SetRowHeight(1, 30);
sheet.SetRowHidden(3, hidden: true);
sheet.SetColumnHidden(4, hidden: true);

workbook.Save("resumen.xlsx");
```

## Notas

- El ancho de una columna se mide en **caracteres** de la fuente por defecto (como en Excel) y el
  alto de una fila, en **puntos**.
- Ocultar una fila o una columna **no borra su contenido**: sigue ahí y reaparece al mostrarla.
- Un panel **dividido** (no congelado) no se modela: se conserva al guardar.