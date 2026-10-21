---
id: combinacion-de-celdas
title: Combinación de celdas
sidebar_label: Combinación de celdas
---

# Combinación de celdas

Combinar celdas es una operación con dos reglas: la combinación **horizontal** mueve el contenido de
las celdas que absorbe, y la **vertical** solo escribe las marcas.

## API

| Miembro | Qué hace |
| --- | --- |
| `TableRow.MergeCells(int startIndex, int count)` | Combina celdas seguidas en horizontal. |
| `TableRow.UnmergeCell(int cellIndex)` | Deshace la combinación de una celda. |
| `Table.MergeCellsVertically(int firstRowIndex, int cellIndex, int rowCount)` | Combina celdas en vertical. |
| `Table.UnmergeCellsVertically(int firstRowIndex, int cellIndex, int rowCount)` | Deshace la combinación vertical. |
| `TableCell.ColumnSpan` | Cuántas columnas ocupa la celda. |
| `TableCell.VerticalMerge` | El estado de la combinación vertical. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Table table = document.AddTable(2, 2);
table.Rows[0].Cells[0].Text = "Cabecera";
table.Rows[0].Cells[1].Text = "Importe";
table.Rows[1].Cells[0].Text = "Total";
table.Rows[1].Cells[1].Text = "1.234,56";

// Horizontal: el contenido de la segunda celda pasa a la primera.
table.Rows[0].MergeCells(0, 2);

// Vertical: solo se escriben las marcas.
table.MergeCellsVertically(0, 0, 2);

document.Save("informe.docx");
```

## Notas

- Una combinación **rectangular** (una celda que participa en dos combinaciones verticales) se rechaza con
  `ArgumentException` en vez de producir una tabla rota.
- La combinación horizontal **mueve** el contenido: no se pierde ningún párrafo ni ninguna tabla
  anidada (ADR 0015).