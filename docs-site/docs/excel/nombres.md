---
id: nombres
title: Nombres
sidebar_label: Nombres
---

# Nombres

Un nombre apunta a un rango y se puede usar en las fórmulas. Un nombre tiene **ámbito de libro** o
**ámbito de hoja**.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.AddNamedRange(string name, string rangeReference)` | Crea un nombre de ámbito de libro. |
| `Worksheet.AddLocalNamedRange(string name, string rangeReference)` | Crea un nombre de ámbito de esa hoja. |
| `Worksheet.RemoveLocalNamedRange(string name)` | Elimina un nombre de esa hoja. |
| `ExcelWorkbook.DefinedNames` | Los nombres del libro. |
| `ExcelWorkbook.FindDefinedName(string name)` | Busca un nombre. |
| `ExcelWorkbook.RemoveNamedRange(string name)` | Elimina un nombre de ámbito de libro. |
| `ExcelDefinedName.Name`, `RefersTo`, `Scope`, `IsRange`, `Range` | Los datos del nombre. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet["A1"].Value = "Enero";
sheet["A2"].Value = "Febrero";
sheet["B1"].Value = 10d;
sheet["B2"].Value = 20d;

sheet.AddNamedRange("Meses", "A1:A2");
sheet.AddLocalNamedRange("Importes", "B1:B2");

ExcelDefinedName? name = workbook.FindDefinedName("meses");
Console.WriteLine($"{name?.Name} -> {name?.RefersTo}");

workbook.Save("resumen.xlsx");
```

## Notas

- El nombre se guarda en forma absoluta (`Hoja1!$A$1:$A$2`), como hace Excel.
- Un nombre que ya existe **se sustituye** si se vuelve a crear; no puede haber nombres duplicados
  en el mismo ámbito.
- Los nombres que Excel crea por su cuenta (las áreas de impresión, por ejemplo) se listan y se
  conservan, pero no se interpretan.