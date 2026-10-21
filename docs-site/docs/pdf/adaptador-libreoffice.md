---
id: adaptador-libreoffice
title: Adaptador de LibreOffice
sidebar_label: Adaptador de LibreOffice
---

# Adaptador de LibreOffice

El adaptador lanza LibreOffice en modo **headless** y lo vigila: cada conversión usa su propio
directorio temporal y su propio perfil de usuario, para no tocar el LibreOffice que el usuario
tenga abierto.

## API

| Miembro | Qué hace |
| --- | --- |
| `LibreOfficePdfConverter()` | Crea el conversor con la configuración por omisión. |
| `LibreOfficePdfConverterOptions.ExecutablePath` | Dónde está el ejecutable, si no se busca solo. |
| `LibreOfficePdfConverterOptions.MinimumVersion` | La versión mínima aceptada (7.0 por omisión). |
| `LibreOfficePdfConverterOptions.Timeout` | Cuánto puede tardar la conversión. |
| `LibreOfficePdfConverterOptions.MaximumConcurrentConversions` | Cuántas conversiones a la vez. |
| `LibreOfficePdfConverter.IsAvailable` | Si el ejecutable está. |
| `LibreOfficePdfConverter.ExecutablePath`, `Version` | Dónde está y qué versión tiene. |

## Ejemplo

```csharp
using OfficeEngine.Pdf.LibreOffice;

using LibreOfficePdfConverter converter = new();

if (!converter.IsAvailable)
{
    Console.WriteLine("LibreOffice no está instalado.");
    return;
}

Console.WriteLine($"LibreOffice {converter.Version} en {converter.ExecutablePath}");
```

## Qué hace el adaptador por debajo

- Busca el ejecutable en las rutas habituales y en el `PATH`. En Windows usa el compañero de
  consola (`soffice.com`), porque el gráfico se desvincula y no devuelve el código de salida.
- Comprueba la **versión** antes de convertir, y rechaza las anteriores a la mínima.
- Aísla la conversión: directorio temporal propio, perfil de usuario propio y archivos que se
  borran al terminar.
- Mira el **código de salida** y el **archivo producido**: el PDF tiene que existir, no estar vacío y
  empezar por `%PDF-`.

## Notas

- En Windows hace falta `soffice.com`, que viene con la instalación de LibreOffice.
- El motor escribe ruido en la salida de error (avisos del proveedor de scripts): es información,
  no un error.