---
id: graficos
title: Gráficos
sidebar_label: Gráficos
---

# Gráficos

Un gráfico se ancla en una celda y apunta a los rangos de sus categorías y de sus valores. Además
copia esos valores, para que se pueda mostrar sin recalcular.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.AddChart(string cellReference, ExcelChartType type, string categoriesRange, string valuesRange, string? title = null, string? seriesName = null, double widthPoints = 480, double heightPoints = 288)` | Crea el gráfico. |
| `Worksheet.Charts` | Los gráficos de la hoja. |
| `Worksheet.RemoveChart(string name)` | Elimina el gráfico. |
| `ExcelChart.Name`, `Reference`, `Type`, `CategoriesRange`, `ValuesRange`, `Title`, `SeriesName` | Los datos del gráfico. |

Los tipos son `ExcelChartType.Column`, `Bar`, `Line` y `Pie`.

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet["A1"].Value = "Enero";
sheet["A2"].Value = "Febrero";
sheet["A3"].Value = "Marzo";
sheet["B1"].Value = 10d;
sheet["B2"].Value = 20d;
sheet["B3"].Value = 15d;

sheet.AddChart(
    "D2",
    ExcelChartType.Column,
    categoriesRange: "A1:A3",
    valuesRange: "B1:B3",
    title: "Ventas",
    seriesName: "Total");

workbook.Save("resumen.xlsx");
```

## Notas

- El gráfico tiene **una sola serie**: los gráficos de varias series o de áreas combinadas se
  conservan al guardar, pero no se crean.
- La caché de datos es la foto del momento de crear el gráfico; no se refresca sola.