---
id: validacion-de-datos
title: Validación de datos
sidebar_label: Validación de datos
---

# Validación de datos

Una regla de validación dice qué se puede escribir en un rango. La biblioteca **escribe** la regla;
la comprueba Excel, no la librería.

## API

| Miembro | Qué hace |
| --- | --- |
| `Worksheet.AddDataValidation(ExcelDataValidation validation)` | Añade una regla a la hoja. |
| `ExcelDataValidation.List(string rangeReference, IEnumerable<string> values, bool allowBlank = true)` | Una lista de valores permitidos. |
| `ExcelDataValidation.WholeNumber(string rangeReference, double? minimum = null, double? maximum = null, bool allowBlank = true)` | Un número entero dentro de unos límites. |
| `ExcelDataValidation.DecimalNumber(string rangeReference, double? minimum = null, double? maximum = null, bool allowBlank = true)` | Un número con decimales dentro de unos límites. |
| `Worksheet.DataValidations` | Las reglas de la hoja. |
| `ExcelDataValidation.Range`, `Type`, `AllowedValues`, `Minimum`, `Maximum`, `AllowBlank` | Los datos de la regla. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet.AddDataValidation(ExcelDataValidation.List("C1:C10", ["bajo", "medio", "alto"]));
sheet.AddDataValidation(ExcelDataValidation.WholeNumber("D1:D10", minimum: 1, maximum: 100));

workbook.Save("resumen.xlsx");
```

## Notas

- La lista se guarda **en línea**, con los valores separados por comas: un valor con coma o comilla
  doble se rechaza con `ArgumentException`.
- Las reglas de tipo fecha, hora, longitud o fórmula personalizada **se leen y se conservan**, pero
  esta versión no las crea.
- Una regla que se aplica a varias áreas se informa por su primera área; el resto se conserva.