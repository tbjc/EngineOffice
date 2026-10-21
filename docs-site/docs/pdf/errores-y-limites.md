---
id: errores-y-limites
title: Errores y límites
sidebar_label: Errores y límites
---

# Errores y límites

La conversión falla con un error **tipado**, y siempre con su motivo.

## Errores

| Situación | Error |
| --- | --- |
| El motor no está instalado, o su versión es anterior a la mínima | `PdfConversionException` |
| El motor termina con un código de salida distinto de cero | `PdfConversionException` (con el código) |
| El motor termina sin producir el PDF | `PdfConversionException` |
| El archivo producido no es un PDF | `PdfConversionException` |
| El motor no termina a tiempo | `PdfConversionException` (se mata el proceso) |
| El documento de entrada está vacío o no se puede leer | `ArgumentException` |
| El documento de entrada supera el tamaño máximo | `OfficeResourceLimitExceededException` |
| El PDF produced supera el tamaño máximo | `OfficeResourceLimitExceededException` |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Pdf;
using OfficeEngine.Pdf.LibreOffice;

using LibreOfficePdfConverter converter = new();

try
{
    using FileStream file = File.OpenRead("informe.docx");
    PdfConversionResult result = converter.Convert(file, PdfSourceFormat.WordDocument);
    File.WriteAllBytes("informe.pdf", result.Pdf);
}
catch (PdfConversionException error)
{
    Console.WriteLine($"La conversión falló: {error.Message}");
    foreach (string message in error.Diagnostics)
    {
        Console.WriteLine($"  {message}");
    }
}
```

## Los límites, por omisión

| Límite | Valor |
| --- | --- |
| Tamaño máximo de la entrada | 64 MiB |
| Tamaño máximo del PDF | 64 MiB |
| Tiempo máximo de la conversión | 2 minutos |
| Versión mínima de LibreOffice | 7.0 |

## Notas

- Los **diagnósticos** que devuelve la conversión describen la conversión, nunca el contenido del
  documento.
- El mensaje de error **no** incluye el texto del documento; el detalle del motor queda en la
  excepción interna, que es donde se puede pedir de forma explícita.