---
id: celdas
title: Celdas
sidebar_label: Celdas
---

# Celdas

Una celda se direcciona por su referencia A1 (`A1`) o por sus índices **1-based** de fila y
columna (`sheet[3, 2]` es `B3`).

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.GetCell(string cellReference)` | La celda por su referencia. |
| `Worksheet.GetCell(int row, int column)` | La celda por sus índices 1-based. |
| `Worksheet["B2"]`, `Worksheet[3, 2]` | Atajos de lo anterior. |
| `Cell.Value` | El valor: `string`, `double`, `bool`, `DateTime` o `null`. |
| `Cell.Reference`, `Cell.Row`, `Cell.Column` | Dónde está la celda. |
| `Cell.SetDateTime(DateTime value, string format)` | Escribe una fecha con su formato. |
| `Cell.NumberFormat` | El código de formato efectivo. |
| `Cell.Style` | El estilo de la celda. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet["A1"].Value = "Concepto";
sheet["B1"].Value = 1250.75;
sheet["C1"].Value = true;
sheet["D1"].Value = "sin valor";

Worksheet other = workbook.Worksheets[0];
Cell cell = other.GetCell(1, 2);
Console.WriteLine($"{cell.Reference}: {cell.Value}");

workbook.Save("resumen.xlsx");
```

## Notas

- Asignar `null` **limpia** la celda; no crea una celda vacía.
- Un `DateTime` necesita un formato de fecha: si la celda no lo tiene, `Value` lanza
  `InvalidOperationException` en vez de inventar uno.
- Leer una celda que no existe devuelve `null` y **no la crea**.