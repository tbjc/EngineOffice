---
id: imagenes
title: Imágenes
sidebar_label: Imágenes
---

# Imágenes

Una imagen se incrusta **en línea**: ocupa su sitio en el flujo del texto. Cada imagen crea su propia
parte y su relación.

## API

| Miembro | Qué hace |
| --- | --- |
| `Paragraph.AddImage(string path, double widthPoints, double heightPoints)` | Incrusta una imagen de un archivo. |
| `Paragraph.AddImage(Stream stream, ImageFormat format, double widthPoints, double heightPoints)` | Incrusta una imagen de un flujo. |
| `Paragraph.Images` | Las imágenes incrustadas del párrafo. |
| `Image.WidthPoints`, `Image.HeightPoints` | El tamaño con el que se insertó. |
| `Image.ContentType` | El tipo de contenido de la parte. |
| `Image.ReadAllBytes()` | Los bytes de la imagen. |

Los formatos admitidos son `ImageFormat.Png`, `Jpeg`, `Gif`, `Bmp`, `Tiff`, `Emf` y `Wmf`.

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Paragraph figure = document.AddParagraph("Figura 1");
figure.AddImage("logo.png", 120.0, 60.0);

// También desde un flujo, si la imagen no está en un archivo.
using FileStream stream = File.OpenRead("logo.png");
figure.AddImage(stream, ImageFormat.Png, 120.0, 60.0);

document.Save("informe.docx");
```

## Notas

- El tamaño se indica en **puntos** y se convierte a EMU al escribir. La biblioteca no reescala la
  imagen ni la recorta.
- Las imágenes **flotantes** (`wp:anchor`), las **vinculadas** y las que llevan texto alternativo
  no se modelan: se conservan al guardar.
- Funciona en el cuerpo, en las celdas de tabla y en los encabezados y pies.