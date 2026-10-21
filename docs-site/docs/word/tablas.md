---
id: tablas
title: Tablas
sidebar_label: Tablas
---

# Tablas

Una tabla tiene filas, celdas y un conjunto de propiedades. Cada celda contiene párrafos.

## API

| Miembro | Qué hace |
| --- | --- |
| `WordDocument.AddTable(int rows, int columns)` | Añade una tabla al final del cuerpo. |
| `Table.Rows`, `RowCount`, `ColumnCount` | Las filas y las medidas. |
| `Table.AddRow()` | Añade una fila al final. |
| `Table.Width`, `Alignment`, `Layout` | Las propiedades de la tabla. |
| `Table.ShadingColor` | El sombreado de la tabla. |
| `Table.BorderTop`, `BorderLeft`, `BorderBottom`, `BorderRight`, `BorderInsideHorizontal`, `BorderInsideVertical` | Los bordes. |
| `TableRow.Cells`, `CellCount` | Las celdas de la fila. |
| `TableCell.Paragraphs`, `Text` | El contenido de la celda. |
| `TableCell.AddParagraph(string text)` | Añade un párrafo a la celda. |
| `TableCell.Width`, `VerticalAlignment`, `ShadingColor` | Las propiedades de la celda. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

Table totals = document.AddTable(2, 2);
totals.Rows[0].Cells[0].Text = "Concepto";
totals.Rows[0].Cells[1].Text = "Importe";
totals.Rows[1].Cells[0].Text = "Total";
totals.Rows[1].Cells[1].Text = "1.234,56";

totals.Width = TableWidth.Percentage(100);
totals.BorderBottom = new TableBorder(BorderStyle.SingleLine, 1.0);

document.Save("informe.docx");
```

## Notas

- `TableCell.Text` **reemplaza** el contenido de la celda por un único párrafo. Para varias líneas,
  usa `AddParagraph`.
- Las tablas **anidadas** dentro de una celda se conservan al guardar, pero no se modelan.