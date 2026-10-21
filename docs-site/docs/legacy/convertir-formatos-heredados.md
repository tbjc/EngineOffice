---
id: convertir-formatos-heredados
title: Convertir formatos heredados
sidebar_label: Convertir formatos heredados
---

# Convertir formatos heredados

Los formatos antiguos (`.doc` de Word 97-2003 y `.xls` de Excel 97-2003) no se leen: **se
convierten** con el motor externo. La diferencia entre convertir, leer, editar y guardar en el
formato original es la diferencia entre una capacidad y otra.

## API

| Miembro | Qué hace |
| --- | --- |
| `LegacyConverter.Convert(Stream source, LegacyDocumentFormat format)` | Convierte el documento. |
| `LegacyDocumentFormat.Word97To2003`, `Excel97To2003` | El formato de partida. |
| `LegacyConverterOptions.ExecutablePath`, `MinimumVersion`, `Timeout`, `MaximumInputBytes` | La configuración. |
| `LegacyConverter.IsAvailable`, `Version` | Si el motor está y qué versión tiene. |
| `LegacyConversionResult.Document` | Los bytes del documento convertido. |
| `LegacyConversionResult.Diagnostics`, `Duration` | Lo que dijo el motor y cuánto tardó. |

## Ejemplo

```csharp
using OfficeEngine.Legacy;
using OfficeEngine.Word;

using FileStream legacy = File.OpenRead("antiguo.doc");

LegacyConverter converter = new();
LegacyConversionResult result = converter.Convert(legacy, LegacyDocumentFormat.Word97To2003);
File.WriteAllBytes("antiguo.docx", result.Document);

// A partir de aquí, el documento se trabaja con la API normal.
using MemoryStream converted = new(result.Document);
using WordDocument document = WordDocument.Open(converted);
Console.WriteLine(document.Text);
```

## Qué se pierde

La conversión es la interpretación que hace el motor del original, no una traducción byte a byte. No
se conservan:

- las **macros**;
- los **objetos incrustados** y los **vínculos OLE**;
- la **maquetación exacta**;
- los **metadatos** del original.

## Lo que **no** existe

- Leer un `.doc` o un `.xls` directamente.
- Editar un archivo heredado en su formato.
- Guardar en el formato heredado: guardar escribe el formato moderno.

## Notas

- La conversión necesita el motor instalado; sin él, `Convert` lanza `LegacyConversionException`.
- Cada conversión arranca un proceso del motor: cuesta unos segundos.