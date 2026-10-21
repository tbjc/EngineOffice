---
id: crear-y-guardar
title: Crear y guardar
sidebar_label: Crear y guardar
---

# Crear y guardar

Un documento se crea en memoria, se edita y se guarda en un archivo o en un flujo.

## API

| Miembro | Qué hace |
| --- | --- |
| `WordDocument.Create()` | Crea un documento vacío con una sección US Letter vertical y márgenes de una pulgada. |
| `WordDocument.Open(string path)` | Abre un documento de un archivo. |
| `WordDocument.Open(Stream stream, bool leaveOpen = false)` | Abre un documento de un flujo, copiándolo a memoria. |
| `Save(string path)` | Guarda en un archivo, de forma atómica (escribe en un temporal y lo mueve). |
| `Save(Stream destination)` | Guarda en un flujo, que queda abierto. |
| `Dispose()` | Libera el paquete. Un documento liberado lanza `ObjectDisposedException`. |
| `Paragraphs` | Los párrafos de primer nivel del cuerpo. |
| `Tables` | Las tablas de primer nivel del cuerpo. |
| `Sections` | Las secciones del documento. |
| `Text` | El texto visible del documento. |
| `AddParagraph(string text)` | Añade un párrafo al final del cuerpo y devuelve una vista viva de él. |
| `AddTable(int rows, int columns)` | Añade una tabla al final del cuerpo. |
| `AddSection(SectionStartType startType)` | Añade una sección al final. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

document.AddParagraph("Informe mensual");
document.AddParagraph("Generado con OfficeEngine");

document.Save("informe.docx");

// Volver a abrirlo y leer lo que se escribió.
using WordDocument reopened = WordDocument.Open("informe.docx");
Console.WriteLine(reopened.Text);
```

## Notas

- El contenido del cuerpo que no modela la biblioteca (un control de contenido, por ejemplo) **se
  conserva** al guardar y se avisa por `document.Diagnostics`.
- `Save` no deja el archivo a medias: si falla, el destino anterior sigue intacto.