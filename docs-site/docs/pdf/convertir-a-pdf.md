---
id: convertir-a-pdf
title: Convertir a PDF
sidebar_label: Convertir a PDF
---

# Convertir a PDF

La conversión a PDF la hace un **motor externo**, no la biblioteca. El contrato vive en
`OfficeEngine.Pdf` y el motor, en un adaptador aparte.

## API

| Miembro | Qué hace |
| --- | --- |
| `IPdfConverter.Convert(Stream source, PdfSourceFormat format, PdfConversionOptions? options = null)` | Convierte un documento a PDF. |
| `PdfSourceFormat.WordDocument`, `ExcelWorkbook` | El tipo de documento de partida. |
| `PdfConversionOptions.MaximumInputBytes`, `MaximumOutputBytes` | Los límites de tamaño. |
| `PdfConversionOptions.Timeout` | Cuánto puede tardar. |
| `PdfConversionOptions.KeepTemporaryFiles` | Conserva los temporales, para diagnosticar. |
| `PdfConversionResult.Pdf` | Los bytes del PDF. |
| `PdfConversionResult.Diagnostics` | Lo que dijo el motor. |
| `PdfConversionResult.Duration` | Cuánto tardó. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Pdf;
using OfficeEngine.Pdf.LibreOffice;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
workbook.Worksheets[0]["A1"].Value = "Informe para imprimir";
using MemoryStream current = new();
workbook.Save(current);
current.Position = 0;

using LibreOfficePdfConverter converter = new();
PdfConversionResult result = converter.Convert(current, PdfSourceFormat.ExcelWorkbook);
File.WriteAllBytes("informe.pdf", result.Pdf);
```

## Notas

- La conversión **parte de los bytes que se le dan**: para exportar el estado actual de un documento,
  guárdalo en memoria y entrega ese flujo. Nunca se reutiliza una copia antigua en disco.
- El motor externo **no se distribuye** con la biblioteca: lo instala el usuario.
- El PDF depende de la versión del motor, de las fuentes instaladas y de la configuración del
  sistema: **no hay equivalencia** con lo que imprimiría Microsoft Office.