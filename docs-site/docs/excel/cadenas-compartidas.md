---
id: cadenas-compartidas
title: Cadenas compartidas
sidebar_label: Cadenas compartidas
---

# Cadenas compartidas

El texto de las celdas se guarda en una **tabla de cadenas compartidas**: el mismo texto se escribe
**una sola vez** y cada celda guarda su índice. Es lo que hace Excel, y es lo que hace que un libro
con mucho texto repetido no crezca sin medida.

## API

| Miembro | Qué hace |
| --- | --- |
| `Cell.Value` | Al escribir un texto, se añade a la tabla si no estaba. |
| `ExcelWorkbook.Diagnostics` | Lo que la biblioteca ha encontrado y no modela. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

// El mismo texto en mil celdas: en el archivo hay una sola entrada.
for (int row = 1; row <= 1000; row++)
{
    sheet[row, 1].Value = "Pendiente de revisión";
    sheet[row, 2].Value = "Pendiente de revisión";
}

workbook.Save("resumen.xlsx");
```

## Notas

- La biblioteca **deduplica** el texto: dos celdas con el mismo texto apuntan al mismo índice.
- Los libros ajenos que ya tienen tabla de cadenas compartidas se leen igual, y las celdas con
  `t="s"` se resuelven contra ella.
- Un índice que apunta a una entrada que no existe hace que el archivo se considere mal formado
  (`InvalidOfficeFileException`).