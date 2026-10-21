---
id: formulas
title: Fórmulas
sidebar_label: Fórmulas
---

# Fórmulas

La biblioteca **escribe** fórmulas y, además, **calcula** un subconjunto de ellas con su propio motor.
Escribir una fórmula y calcularla son dos cosas distintas.

## API

| Miembro | Qué hace |
| --- | --- |
| `Cell.Formula` | La fórmula de la celda, con o sin el `=` inicial. |
| `ExcelWorkbook.EvaluateFormula(string formula, string? sheetName = null)` | Evalúa una fórmula sobre los **valores** del libro. |
| `FormulaEngine.Evaluate(string formula, IFormulaContext context)` | El motor, para quien traiga sus propios valores. |
| `FormulaEngine.SupportedFunctions` | Las funciones soportadas. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Formulas;

using ExcelWorkbook workbook = ExcelWorkbook.Create();
Worksheet sheet = workbook.Worksheets[0];

sheet["A1"].Value = 6d;
sheet["A2"].Value = 7d;

// Escribir la fórmula.
sheet["B1"].Formula = "=A1*A2";

// Calcularla con el motor propio.
object? result = workbook.EvaluateFormula("=A1*A2");
Console.WriteLine(result);

// El nombre de las funciones también se acepta en español.
Console.WriteLine(workbook.EvaluateFormula("=SUMA(A1,A2)"));

workbook.Save("resumen.xlsx");
```

## Qué se calcula y qué no

La matriz completa está en [la matriz de fórmulas](matriz-de-formulas.md). En corto:

- Se calculan: números, textos, booleanos, referencias (`A1`, `Hoja1!A1`), rangos (`A1:B2`), los
  operadores aritméticos, de comparación y de concatenación, y 32 funciones (con el nombre en
  inglés y en español).
- **No** se calculan: las funciones volátiles (`HOY`, `AHORA`), las matrices dinámicas, las
  referencias externas ni los nombres definidos. Devuelven `#NAME?` o `#VALUE!`, nunca un valor
  inventado.

## Los nombres en español

El motor acepta el nombre de cada función **en inglés y en español**: `=SUM(A1:A2)` y `=SUMA(A1:A2)`
son la misma fórmula.

Como Excel guarda los nombres de las funciones **en inglés** dentro del archivo (lo que cambia es
lo que muestra la interfaz), la biblioteca **traduce** la fórmula al nombre canónico cuando la
escribe: `=SUMA(A1:A2)` se guarda como `=SUM(A1:A2)`, que es lo que Excel evalúa. El texto de los
literales no se toca.

## Notas

- Al escribir una fórmula, la celda **no** tiene resultado guardado: `Value` devuelve `null` hasta que
  Excel (u otro motor) recalcula el libro.
- El motor **no sigue dependencias**: evalúa sobre los valores que hay en las celdas, no sobre las
  fórmulas que contienen.