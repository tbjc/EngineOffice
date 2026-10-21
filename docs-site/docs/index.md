---
id: index
title: OfficeEngine
sidebar_label: Inicio
---

# OfficeEngine

Biblioteca .NET independiente para **crear, leer, modificar, guardar y convertir** documentos de Word
(`.docx`) y libros de Excel (`.xlsx`), y para exportarlos a PDF con un motor externo.

Todo lo que hace la biblioteca es **explícito**: no adivina, no inventa contenido y no descarta en
silencio lo que no modela. Cuando algo no se puede leer, lanza un error tipado; cuando algo no se
entiende, lo conserva y lo avisa.

## Instalación

| Paquete | Contenido |
| --- | --- |
| `OfficeEngine.Word` | Modelo y edición de documentos `.docx` |
| `OfficeEngine.Excel` | Modelo y edición de libros `.xlsx` |
| `OfficeEngine.Pdf` | Contratos de conversión a PDF |
| `OfficeEngine.Pdf.LibreOffice` | Adaptador de conversión con LibreOffice headless |
| `OfficeEngine.Legacy` | Conversión de los formatos heredados `.doc` y `.xls` |
| `OfficeEngine.Formulas` | Motor de cálculo de fórmulas |

## Un primer ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Paragraph title = document.AddParagraph("Informe mensual");
title.Alignment = ParagraphAlignment.Center;

Run amount = title.AddRun("Total: 1.234,56");
amount.Bold = true;
amount.FontSizePoints = 12.0;

document.Save("informe.docx");
```

## Qué hay en esta documentación

| Sección | Qué cubre |
| --- | --- |
| [Word](word/crear-y-guardar.md) | Documentos de Word, del párrafo al encabezado, con las tablas, las imágenes, los hipervínculos y la búsqueda. |
| [Excel](excel/crear-y-guardar.md) | Libros de Excel: celdas, rangos, estilos, fórmulas, nombres, validaciones, filtros, tablas, gráficos e impresión. |
| [PDF](pdf/convertir-a-pdf.md) | Conversión a PDF con un motor externo, detrás de una interfaz. |
| [Formatos heredados](legacy/convertir-formatos-heredados.md) | Conversión de `.doc` y `.xls` a los formatos modernos. |

## Cómo está pensado

- **Todo en memoria.** El paquete se abre, se edita y se guarda; mientras está abierto, vive en
  memoria. No hay bloqueos de archivo y el guardado es atómico.
- **Síncrono.** No hay métodos `async`: el trabajo es de memoria, no de espera de E/S.
- **Errores tipados.** Un archivo ilegible produce `InvalidOfficeFileException`; un límite superado,
  `OfficeResourceLimitExceededException`; una petición no soportada,
  `UnsupportedOfficeFeatureException`.
- **Lo que no se modela, se conserva.** Un elemento desconocido sobrevive al guardado y, cuando
  importa, se avisa por diagnósticos.