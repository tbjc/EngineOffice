---
id: crear-y-guardar
title: Crear y guardar
sidebar_label: Crear y guardar
---

# Crear y guardar

Un libro se crea en memoria, se edita y se guarda. Las hojas se añaden, se renombran y se eliminan.

## API

| Miembro | Qué hace |
| --- | --- |
| `ExcelWorkbook.Create()` | Crea un libro con una hoja llamada `Sheet1`. |
| `ExcelWorkbook.Open(string path)` | Abre un libro de un archivo. |
| `ExcelWorkbook.Open(Stream stream, bool leaveOpen = false)` | Abre un libro de un flujo. |
| `Save(string path)` / `Save(Stream destination)` | Guarda, de forma atómica en el caso del archivo. |
| `Dispose()` | Libera el paquete. |
| `Workbook.Worksheets` | Las hojas del libro. |
| `Worksheets.Add(string name)` | Añade una hoja al final. |
| `Worksheets.Remove(string name)` | Elimina una hoja por su nombre. |
| `Worksheets["Resumen"]`, `Worksheets[0]` | Una hoja por nombre o por posición. |
| `Workbook.Diagnostics` | Lo que se ha encontrado y no se modela. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();

Worksheet sheet = workbook.Worksheets.Add("Resumen");
sheet["A1"].Value = "Concepto";
sheet["B1"].Value = "Importe";

workbook.Save("resumen.xlsx");

using ExcelWorkbook reopened = ExcelWorkbook.Open("resumen.xlsx");
Console.WriteLine(reopened.Worksheets["Resumen"]["A1"].Value);
```

## Notas

- Un libro no puede quedarse sin hojas: `Worksheets.Remove` lanza `InvalidOperationException` si
  queda una sola.
- El nombre de una hoja no puede repetirse, ni llevar `\ / ? * [ ] :`, ni pasar de 31 caracteres.