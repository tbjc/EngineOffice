---
id: imagenes
title: Imágenes
sidebar_label: Imágenes
---

# Imágenes

Una imagen se ancla en una celda y crea su propia parte de medios.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.AddImage(string cellReference, Stream content, ExcelImageFormat format, double widthPoints, double heightPoints, string? name = null)` | Ancla una imagen. |
| `Worksheet.Images` | Las imágenes de la hoja. |
| `Worksheet.RemoveImage(string name)` | Elimina la imagen. |
| `ExcelImage.Name`, `Reference`, `WidthPoints`, `HeightPoints`, `MediaType` | Los datos de la imagen. |

Los formatos son `ExcelImageFormat.Png`, `Jpeg`, `Gif`, `Bmp` y `Tiff`.

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

using FileStream stream = File.OpenRead("logo.png");
sheet.AddImage("A1", stream, ExcelImageFormat.Png, 120.0, 60.0, name: "Logo");

workbook.Save("resumen.xlsx");
```

## Notas

- El tamaño se indica en **puntos** y no se recalcula con el tamaño de las celdas.
- La biblioteca no valida que los bytes sean realmente una imagen del formato declarado.