---
id: fechas-y-formatos-numericos
title: Fechas y formatos numéricos
sidebar_label: Fechas y formatos numéricos
---

# Fechas y formatos numéricos

En Excel una fecha **es un número**: el serial. La biblioteca escribe el serial y el formato que
dice que se lea como fecha.

## API

| Miembro | Qué hace |
| --- | --- |
| `Cell.SetDateTime(DateTime value, string format)` | Escribe la fecha y su formato en una sola llamada. |
| `Cell.NumberFormat` | El código de formato efectivo (`General` si no hay ninguno). |
| `Cell.Value` | Devuelve un `DateTime` cuando la celda tiene un formato de fecha. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet["A1"].Value = "Fecha de cierre";
sheet["B1"].SetDateTime(new DateTime(2026, 1, 31), "yyyy-mm-dd");
sheet["B2"].SetDateTime(new DateTime(2026, 1, 31), "dd/mm/yyyy hh:mm");

// El formato también se puede cambiar por separado.
sheet["B3"].NumberFormat = "#,##0.00";
sheet["B3"].Value = 1234.5;

workbook.Save("resumen.xlsx");

using ExcelWorkbook reopened = ExcelWorkbook.Open("resumen.xlsx");
Console.WriteLine(reopened.Worksheets[0]["B1"].Value);
```

## Notas

- Asignar un `DateTime` a `Value` **sin** un formato de fecha lanza `InvalidOperationException`: la
  biblioteca no adivina el formato.
- Los formatos integrados que dependen de la configuración regional (los identificadores 27 a 36 y 50
  a 58) se reconocen como fecha, pero su código se devuelve como `null`, porque no hay un único
  código universal.
- Las fechas se guardan con el sistema de fechas del libro (1900 o 1904), y se leen con el mismo.