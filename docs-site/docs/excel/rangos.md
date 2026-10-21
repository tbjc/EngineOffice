---
id: rangos
title: Rangos
sidebar_label: Rangos
---

# Rangos

Un rango se direcciona por su referencia (`A1:B2`) o por sus esquinas, en índices 1-based. Es una
**vista viva**: leerlo o escribirlo actúa sobre el libro en ese momento.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.GetRange(string rangeReference)` | El rango por su referencia. |
| `Worksheet.GetRange(int firstRow, int firstColumn, int lastRow, int lastColumn)` | El rango por sus esquinas. |
| `CellRange.Reference` | La referencia del rango, en forma A1. |
| `CellRange.FirstRow`, `FirstColumn`, `LastRow`, `LastColumn`, `RowCount`, `ColumnCount` | Sus medidas. |
| `CellRange.Cells` | Las celdas, en orden de fila. |
| `CellRange.GetValues()` | Los valores, en una matriz con índices 0-based desde la esquina. |
| `CellRange.SetValues(object?[,] values)` | Escribe una matriz de valores. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet.GetRange("A1:B3").SetValues(new object?[,]
{
    { "Concepto", "Importe" },
    { "Ventas", 1250.75 },
    { "Total", 1250.75 },
});

object?[,] values = sheet.GetRange("A1:B3").GetValues();
Console.WriteLine(values[2, 0]);

workbook.Save("resumen.xlsx");
```

## Notas

- `SetValues` exige **exactamente** las medidas del rango y valida todos los valores antes de
  escribir el primero: un tipo no soportado deja el rango intacto.
- Un `null` en la matriz **limpia** la celda y no la crea.
- Un `DateTime` no se puede escribir con `SetValues` (cada celda necesita su propio formato): usa
  `Cell.SetDateTime`.