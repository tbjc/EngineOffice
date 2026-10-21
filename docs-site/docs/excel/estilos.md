---
id: estilos
title: Estilos
sidebar_label: Estilos
---

# Estilos

Un estilo de celda agrupa el formato de carácter, el relleno, los bordes y la alineación. Es
**reutilizable**: el mismo estilo aplicado a varias celdas comparte una entrada del archivo.

## API

| Miembro | Qué hace |
| --- | --- |
| `Cell.Style` | El estilo de la celda. |
| `CellStyle.FontName`, `FontSizePoints`, `Bold`, `Italic`, `Underline`, `FontColor` | El formato de carácter. |
| `CellStyle.FillColor` | El relleno sólido. |
| `CellStyle.Borders` | Los bordes, con `CellBorders` y `CellBorder`. |
| `CellStyle.HorizontalAlignment`, `VerticalAlignment`, `WrapText` | La alineación. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

CellStyle header = new()
{
    Bold = true,
    FontColor = RgbColor.FromRgb(0xFF, 0xFF, 0xFF),
    FillColor = RgbColor.FromRgb(0x1F, 0x49, 0x7D),
    Borders = new CellBorders(bottom: new CellBorder(CellBorderStyle.Thin)),
    HorizontalAlignment = CellHorizontalAlignment.Center,
};

sheet["A1"].Style = header;
sheet["B1"].Style = header;

workbook.Save("resumen.xlsx");
```

## Notas

- El estilo **no** incluye el formato numérico: ese se cambia con `Cell.NumberFormat`, y se conserva
  al aplicar un estilo.
- Cada propiedad es **tri-estado**: `null` significa «el valor por defecto del libro» y `false` desactiva
  la propiedad de forma explícita.
- Aplicar un estilo **reemplaza** las propiedades modeladas; el relleno con patrón o con color de
  tema no se modela y se pierde al aplicarlo.