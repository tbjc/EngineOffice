using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over a single row (<c>w:tr</c>) of a <see cref="Table"/>.
/// </summary>
/// <remarks>
/// This is a lightweight wrapper: it does not own the underlying content and must not be used after
/// the <see cref="WordDocument"/> that produced it has been disposed.
/// </remarks>
public sealed class TableRow
{
    private readonly W.TableRow _element;

    internal TableRow(W.TableRow element) => _element = element;

    /// <summary>Gets a live snapshot of the row's cells, in document order.</summary>
    public IReadOnlyList<TableCell> Cells
        => _element.Elements<W.TableCell>().Select(element => new TableCell(element)).ToArray();

    /// <summary>
    /// Gets the number of cells in the row.
    /// </summary>
    /// <remarks>
    /// A row of a table with merged cells holds fewer cells than the table has columns; the merged
    /// cell carries the columns it covers in <see cref="TableCell.ColumnSpan"/>.
    /// </remarks>
    public int CellCount => _element.Elements<W.TableCell>().Count();

    /// <summary>
    /// Merges two or more adjacent cells of the row into the first one of them.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The first cell of the range keeps its own cell properties, gains the columns the others covered
    /// (<see cref="TableCell.ColumnSpan"/>) and receives their content, in order, after its own; the
    /// other cells are removed. Nothing is discarded: the content of a cell that is merged away is
    /// moved, including a nested table or any other block a foreign document put there.
    /// </para>
    /// <para>
    /// The merged cell states the sum of the widths of the cells it covers when every one of them
    /// states a width of the same kind; otherwise it states no width and the table grid governs the
    /// layout.
    /// </para>
    /// </remarks>
    /// <param name="startIndex">The index of the first cell to merge.</param>
    /// <param name="count">The number of adjacent cells to merge. Must be two or more.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="count"/> is less than two, or the range does not fit in the row.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    /// A cell in the range is part of a vertical merge; merging it horizontally would break that
    /// merge.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored span or width cannot be read.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// A cell in the range stores a width that Office Open XML defines but this version does not model.
    /// </exception>
    public void MergeCells(int startIndex, int count)
    {
        if (startIndex < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(startIndex),
                startIndex,
                "The first cell index must not be negative.");
        }

        if (count < 2)
        {
            throw new ArgumentOutOfRangeException(nameof(count), count, "A merge needs at least two cells.");
        }

        W.TableCell[] cells = _element.Elements<W.TableCell>().ToArray();
        if (startIndex + count > cells.Length)
        {
            throw new ArgumentOutOfRangeException(
                nameof(count),
                count,
                $"The row holds {cells.Length} cells.");
        }

        int span = 0;
        for (int index = startIndex; index < startIndex + count; index++)
        {
            TableCell cell = new(cells[index]);
            if (cell.VerticalMerge is not null)
            {
                throw new InvalidOperationException(
                    "A cell that is part of a vertical merge cannot be merged horizontally.");
            }

            span += cell.ColumnSpan;
        }

        // Read the widths before touching the row: the cells that are merged away must still be there.
        TableWidth? width = MergedWidth(cells, startIndex, count);

        TableCell mergedCell = new(cells[startIndex]);
        mergedCell.ColumnSpan = span;
        mergedCell.Width = width;

        for (int index = startIndex + 1; index < startIndex + count; index++)
        {
            MoveContent(cells[index], cells[startIndex]);
            cells[index].Remove();
        }
    }

    /// <summary>
    /// Splits a cell that covers several columns back into one cell per column.
    /// </summary>
    /// <remarks>
    /// The cell keeps its content and its properties, loses the span and is followed by as many empty
    /// cells as columns it covered. The new cells state no width, so the table grid governs their
    /// layout. Splitting a cell that covers one column only removes a stored span, if there is one.
    /// </remarks>
    /// <param name="cellIndex">The index of the cell to split.</param>
    /// <exception cref="ArgumentOutOfRangeException">The index is not a cell of this row.</exception>
    /// <exception cref="InvalidOperationException">
    /// The cell is part of a vertical merge; splitting it would break that merge.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored span cannot be read.</exception>
    public void UnmergeCell(int cellIndex)
    {
        W.TableCell[] cells = _element.Elements<W.TableCell>().ToArray();
        if (cellIndex < 0 || cellIndex >= cells.Length)
        {
            throw new ArgumentOutOfRangeException(
                nameof(cellIndex),
                cellIndex,
                $"The row holds {cells.Length} cells.");
        }

        TableCell cell = new(cells[cellIndex]);
        if (cell.VerticalMerge is not null)
        {
            throw new InvalidOperationException("A cell that is part of a vertical merge cannot be split.");
        }

        int span = cell.ColumnSpan;
        cell.ColumnSpan = 1;

        OpenXmlElement reference = cells[cellIndex];
        for (int column = 1; column < span; column++)
        {
            // CT_Tc requires content, and an empty paragraph is what a table cell created from scratch
            // carries, so the split cells follow that shape.
            W.TableCell created = new(new W.Paragraph());
            reference.InsertAfterSelf(created);
            reference = created;
        }
    }

    /// <inheritdoc />
    public override string ToString() => $"table row (cells: {CellCount})";

    private static TableWidth? MergedWidth(W.TableCell[] cells, int startIndex, int count)
    {
        TableWidthType? type = null;
        double sum = 0.0;
        for (int index = startIndex; index < startIndex + count; index++)
        {
            if (new TableCell(cells[index]).Width is not { } width || width.Type == TableWidthType.Auto)
            {
                // A width the layout algorithm decides cannot be added up, so the merged cell states
                // none and the table grid gives it its columns' width.
                return null;
            }

            if (type is not null && type != width.Type)
            {
                return null;
            }

            type = width.Type;
            sum += width.Value;
        }

        return type switch
        {
            TableWidthType.Points => TableWidth.Points(sum),
            TableWidthType.Percentage => TableWidth.Percentage(sum),
            _ => null,
        };
    }

    private static void MoveContent(W.TableCell from, W.TableCell to)
    {
        // Everything but w:tcPr moves: the properties belong to the cell that is merged away, while the
        // content must be kept.
        OpenXmlElement[] content = from.ChildElements
            .Where(element => element is not W.TableCellProperties)
            .ToArray();
        foreach (OpenXmlElement child in content)
        {
            child.Remove();
            to.AppendChild(child);
        }
    }
}
