using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over a table (<c>w:tbl</c>) in a <see cref="WordDocument"/>.
/// </summary>
/// <remarks>
/// <para>
/// A table is a block-level element of the document body that sits at the same level as a paragraph,
/// between one paragraph and the next (<c>w:body</c> allows them in any order). A table is made of
/// rows (<c>w:tr</c>), each row is made of cells (<c>w:tc</c>), and each cell holds block-level
/// content — at least the paragraph Word needs to edit it.
/// </para>
/// <para>
/// This version models the structure and the text content of a table, together with the table's own
/// formatting: preferred width, alignment, layout, borders and shading. Cell formatting and cell
/// merge (<c>w:gridSpan</c>/<c>w:vMerge</c>) are modelled on <see cref="TableCell"/>, and the
/// operations that merge and split cells are <see cref="TableRow.MergeCells"/>/<see cref="TableRow.UnmergeCell"/>
/// and the vertical merge methods of this type. Cell margins and the width of the table grid
/// (<c>w:tblGrid</c>) are preserved but not exposed; a table loaded from another tool keeps that
/// markup when its content is edited.
/// </para>
/// <para>
/// This is a lightweight wrapper: it does not own the underlying content and must not be used after
/// the <see cref="WordDocument"/> that produced it has been disposed.
/// </para>
/// </remarks>
public sealed class Table
{
    /// <summary>
    /// Width, in twips, written for every column a table created by <see cref="WordDocument.AddTable"/>
    /// declares. Column widths are not modelled in this version, so the created grid uses this default
    /// instead of measuring the page.
    /// </summary>
    internal const int DefaultColumnWidthTwips = 2400;

    // Child order required by CT_TblPr. The names are the XML local names and the sequence is the
    // one the Open XML SDK generates for a TableProperties element, verified with a reflection probe
    // over DocumentFormat.OpenXml 3.5.1. Properties are inserted following this table instead of
    // being appended, so a table that already carries markup this version does not model (a style,
    // an indent, cell margins) stays valid.
    private static readonly string[] TablePropertyOrder =
    [
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblW", "jc", "tblCellSpacing", "tblInd",
        "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
        "tblPrChange",
    ];

    private readonly W.Table _element;

    internal Table(W.Table element) => _element = element;

    /// <summary>Gets a live snapshot of the table's rows, in document order.</summary>
    public IReadOnlyList<TableRow> Rows
        => _element.Elements<W.TableRow>().Select(element => new TableRow(element)).ToArray();

    /// <summary>Gets the number of rows in the table.</summary>
    public int RowCount => _element.Elements<W.TableRow>().Count();

    /// <summary>
    /// Gets the number of columns the table grid declares (<c>w:tblGrid/w:gridCol</c>).
    /// </summary>
    /// <remarks>
    /// The grid is the schema-authoritative column count. A table whose grid is missing, which the
    /// schema requires but a malformed package can lack, reports <c>0</c>. Merged cells make a row
    /// hold fewer cells than the grid has columns, so this value is not the number of cells in a row.
    /// </remarks>
    public int ColumnCount
        => _element.GetFirstChild<W.TableGrid>()?.Elements<W.GridColumn>().Count() ?? 0;

    /// <summary>
    /// Appends a new row with one empty cell per column of the table grid.
    /// </summary>
    /// <remarks>
    /// The row is appended after the last existing row. The number of cells is
    /// <see cref="ColumnCount"/>; when the table has no grid, the new row has no cells.
    /// </remarks>
    /// <returns>A live view over the new row.</returns>
    public TableRow AddRow()
    {
        W.TableRow row = CreateRow(ColumnCount);
        _element.AppendChild(row);
        return new TableRow(row);
    }

    /// <summary>
    /// Merges the cells that share a column of a run of rows into the topmost one.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The topmost cell of the range is marked as the start of the merge (<c>w:vMerge
    /// w:val="restart"</c>) and the cells below it as its continuation (<c>w:vMerge
    /// w:val="continue"</c>); no cell is created or removed. The cells keep their content: unlike
    /// <see cref="TableRow.MergeCells"/>, a vertical merge moves nothing, so a caller that wants the
    /// merged area to show the text of the cells below has to move it.
    /// </para>
    /// <para>
    /// Every row of the range must hold a cell at <paramref name="cellIndex"/> and all of them must
    /// cover the same number of columns, which is what a rectangular vertical merge needs.
    /// </para>
    /// </remarks>
    /// <param name="firstRowIndex">The index of the top row of the merge.</param>
    /// <param name="cellIndex">The index of the cell, within each row, that is merged.</param>
    /// <param name="rowCount">The number of rows to merge. Must be two or more.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="rowCount"/> is less than two, the range does not fit in the table, or a row of
    /// the range has no cell at <paramref name="cellIndex"/>.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    /// The topmost cell continues a merge that starts in an earlier row, or the cells do not all cover
    /// the same number of columns.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored span or merge cannot be read.</exception>
    public void MergeCellsVertically(int firstRowIndex, int cellIndex, int rowCount)
    {
        TableCell[] cells = ColumnCells(firstRowIndex, cellIndex, rowCount);

        if (cells[0].VerticalMerge == TableVerticalMerge.Continue)
        {
            throw new InvalidOperationException(
                "The topmost cell already continues a vertical merge that starts in an earlier row.");
        }

        int span = cells[0].ColumnSpan;
        foreach (TableCell cell in cells)
        {
            if (cell.ColumnSpan != span)
            {
                throw new InvalidOperationException(
                    "The cells of a vertical merge must all cover the same number of columns.");
            }
        }

        cells[0].VerticalMerge = TableVerticalMerge.Restart;
        for (int row = 1; row < cells.Length; row++)
        {
            cells[row].VerticalMerge = TableVerticalMerge.Continue;
        }
    }

    /// <summary>
    /// Removes the vertical merge from a run of rows.
    /// </summary>
    /// <remarks>
    /// Every cell of the range loses its <c>w:vMerge</c>, whatever part it played in the merge, and
    /// keeps its content. This undoes <see cref="MergeCellsVertically"/>.
    /// </remarks>
    /// <param name="firstRowIndex">The index of the first row of the range.</param>
    /// <param name="cellIndex">The index of the cell, within each row, that is unmerged.</param>
    /// <param name="rowCount">The number of rows. Must be two or more.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="rowCount"/> is less than two, the range does not fit in the table, or a row of
    /// the range has no cell at <paramref name="cellIndex"/>.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored merge cannot be read.</exception>
    public void UnmergeCellsVertically(int firstRowIndex, int cellIndex, int rowCount)
    {
        foreach (TableCell cell in ColumnCells(firstRowIndex, cellIndex, rowCount))
        {
            cell.VerticalMerge = null;
        }
    }

    private TableCell[] ColumnCells(int firstRowIndex, int cellIndex, int rowCount)
    {
        if (firstRowIndex < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(firstRowIndex),
                firstRowIndex,
                "The first row index must not be negative.");
        }

        if (cellIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(cellIndex), cellIndex, "The cell index must not be negative.");
        }

        if (rowCount < 2)
        {
            throw new ArgumentOutOfRangeException(
                nameof(rowCount),
                rowCount,
                "The range must cover at least two rows.");
        }

        IReadOnlyList<TableRow> rows = Rows;
        if (firstRowIndex + rowCount > rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(rowCount), rowCount, $"The table holds {rows.Count} rows.");
        }

        TableCell[] cells = new TableCell[rowCount];
        for (int row = 0; row < rowCount; row++)
        {
            TableRow current = rows[firstRowIndex + row];
            if (cellIndex >= current.CellCount)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(cellIndex),
                    cellIndex,
                    $"The row holds {current.CellCount} cells.");
            }

            cells[row] = current.Cells[cellIndex];
        }

        return cells;
    }

    /// <summary>
    /// Gets or sets the preferred width of the table.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the table does not state a preferred width, which is exactly what
    /// Office Open XML expresses by omitting <c>w:tblW</c>. The widths the table grid declares
    /// (<c>w:tblGrid</c>) are not affected; Word reconciles the two when it lays the table out.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">A stored width cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The table stores a width that Office Open XML defines but this version does not model.
    /// </exception>
    public TableWidth? Width
    {
        get => TableWidth.Read(Properties?.GetFirstChild<W.TableWidth>(), "table width");
        set
        {
            if (value is not { } width)
            {
                Properties?.RemoveAllChildren<W.TableWidth>();
                return;
            }

            width.ApplyTo(EnsureProperty<W.TableWidth>());
        }
    }

    /// <summary>
    /// Gets or sets the alignment of the table between the page margins.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the table inherits its alignment, which is exactly what Office
    /// Open XML expresses by omitting <c>w:jc</c> inside <c>w:tblPr</c>.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is not a defined alignment.</exception>
    /// <exception cref="InvalidOfficeFileException">A stored alignment cannot be read.</exception>
    public TableAlignment? Alignment
    {
        get
        {
            string? stored = Properties?.GetFirstChild<W.TableJustification>()?.Val?.InnerText;
            return stored is null ? null : ToTableAlignment(stored);
        }

        set
        {
            if (value is not { } alignment)
            {
                Properties?.RemoveAllChildren<W.TableJustification>();
                return;
            }

            // Map before touching w:tblPr so that a rejected value leaves no partial markup.
            W.TableRowAlignmentValues ooxml = ToOoxmlAlignment(alignment);
            EnsureProperty<W.TableJustification>().Val = ooxml;
        }
    }

    /// <summary>
    /// Gets or sets how the table lays its columns out.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the table does not state a layout, which is exactly what Office
    /// Open XML expresses by omitting <c>w:tblLayout</c>.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is not a defined layout.</exception>
    /// <exception cref="InvalidOfficeFileException">A stored layout cannot be read.</exception>
    public TableLayout? Layout
    {
        get
        {
            W.TableLayout? stored = Properties?.GetFirstChild<W.TableLayout>();
            if (stored is null)
            {
                return null;
            }

            // An omitted w:type is read as autofit: it is the algorithm the schema assumes and the one
            // Word writes. Only the element being absent at all reads as "no layout".
            return ToTableLayout(stored.Type?.InnerText ?? "autofit");
        }

        set
        {
            if (value is not { } layout)
            {
                Properties?.RemoveAllChildren<W.TableLayout>();
                return;
            }

            W.TableLayoutValues ooxml = ToOoxmlLayout(layout);
            EnsureProperty<W.TableLayout>().Type = ooxml;
        }
    }

    /// <summary>
    /// Gets or sets the shading colour that fills the whole table.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see langword="null"/> means the table states no shading, which is exactly what Office Open
    /// XML expresses by omitting <c>w:shd</c>. A table that carries only a theme fill, which this
    /// version does not model, also reads as <see langword="null"/>.
    /// </para>
    /// <para>
    /// Assigning a colour writes a solid fill (the <c>clear</c> pattern). Assigning
    /// <see langword="null"/> removes the shading this table states.
    /// </para>
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">A stored shading cannot be read.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The table stores a shading pattern that Office Open XML defines but this version does not model.
    /// </exception>
    public RgbColor? ShadingColor
    {
        get => TableShading.Read(Properties?.GetFirstChild<W.Shading>());
        set
        {
            if (value is not { } color)
            {
                Properties?.RemoveAllChildren<W.Shading>();
                return;
            }

            TableShading.Write(EnsureProperty<W.Shading>(), color);
        }
    }

    /// <summary>Gets or sets the top border of the table.</summary>
    /// <remarks>
    /// <see langword="null"/> means the table does not state this edge, so it follows the table style
    /// and the defaults; <c>w:tblBorders</c> is created when the first edge is set and removed again
    /// when the last one is cleared.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">A stored border cannot be read.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The edge uses a line style that Office Open XML defines but this version does not model.
    /// </exception>
    public TableBorder? BorderTop
    {
        get => BorderEdges.Read<W.TopBorder>(Borders);
        set => SetBorder<W.TopBorder>(value);
    }

    /// <summary>Gets or sets the left border of the table.</summary>
    /// <remarks>
    /// See <see cref="BorderTop"/> for the meaning of <see langword="null"/> and for the failures
    /// reading a stored border can raise.
    /// </remarks>
    public TableBorder? BorderLeft
    {
        get => BorderEdges.Read<W.LeftBorder>(Borders);
        set => SetBorder<W.LeftBorder>(value);
    }

    /// <summary>Gets or sets the bottom border of the table.</summary>
    /// <remarks>
    /// See <see cref="BorderTop"/> for the meaning of <see langword="null"/> and for the failures
    /// reading a stored border can raise.
    /// </remarks>
    public TableBorder? BorderBottom
    {
        get => BorderEdges.Read<W.BottomBorder>(Borders);
        set => SetBorder<W.BottomBorder>(value);
    }

    /// <summary>Gets or sets the right border of the table.</summary>
    /// <remarks>
    /// See <see cref="BorderTop"/> for the meaning of <see langword="null"/> and for the failures
    /// reading a stored border can raise.
    /// </remarks>
    public TableBorder? BorderRight
    {
        get => BorderEdges.Read<W.RightBorder>(Borders);
        set => SetBorder<W.RightBorder>(value);
    }

    /// <summary>Gets or sets the border drawn between the rows of the table.</summary>
    /// <remarks>
    /// See <see cref="BorderTop"/> for the meaning of <see langword="null"/> and for the failures
    /// reading a stored border can raise.
    /// </remarks>
    public TableBorder? BorderInsideHorizontal
    {
        get => BorderEdges.Read<W.InsideHorizontalBorder>(Borders);
        set => SetBorder<W.InsideHorizontalBorder>(value);
    }

    /// <summary>Gets or sets the border drawn between the columns of the table.</summary>
    /// <remarks>
    /// See <see cref="BorderTop"/> for the meaning of <see langword="null"/> and for the failures
    /// reading a stored border can raise.
    /// </remarks>
    public TableBorder? BorderInsideVertical
    {
        get => BorderEdges.Read<W.InsideVerticalBorder>(Borders);
        set => SetBorder<W.InsideVerticalBorder>(value);
    }

    /// <inheritdoc />
    public override string ToString()
        => $"table (rows: {RowCount}, columns: {ColumnCount})";

    private W.TableProperties? Properties => _element.GetFirstChild<W.TableProperties>();

    private W.TableBorders? Borders => Properties?.GetFirstChild<W.TableBorders>();

    private W.TableProperties EnsureProperties()
    {
        W.TableProperties? properties = _element.GetFirstChild<W.TableProperties>();
        if (properties is null)
        {
            // CT_Tbl requires w:tblPr before w:tblGrid.
            properties = new W.TableProperties();
            _element.InsertAt(properties, 0);
        }

        return properties;
    }

    private TChild EnsureProperty<TChild>()
        where TChild : OpenXmlElement, new()
    {
        W.TableProperties properties = EnsureProperties();
        TChild? existing = properties.GetFirstChild<TChild>();
        if (existing is not null)
        {
            return existing;
        }

        TChild created = new();
        InsertProperty(properties, created);
        return created;
    }

    private void SetBorder<TEdge>(TableBorder? value)
        where TEdge : W.BorderType, new()
    {
        if (value is not { } border)
        {
            W.TableBorders? borders = Borders;
            if (borders is null)
            {
                return;
            }

            BorderEdges.Write<TEdge>(borders, null);
            if (!borders.HasChildren)
            {
                borders.Remove();
            }

            return;
        }

        BorderEdges.Write<TEdge>(EnsureProperty<W.TableBorders>(), border);
    }

    private static void InsertProperty(W.TableProperties properties, OpenXmlElement element)
        => SchemaOrder.Insert(properties, element, TablePropertyOrder);

    private static TableAlignment ToTableAlignment(string value) => value switch
    {
        "left" => TableAlignment.Left,
        "center" => TableAlignment.Center,
        "right" => TableAlignment.Right,
        _ => throw new InvalidOfficeFileException(
            $"The table stores the alignment '{value}' that this document cannot read."),
    };

    private static W.TableRowAlignmentValues ToOoxmlAlignment(TableAlignment value) => value switch
    {
        TableAlignment.Left => W.TableRowAlignmentValues.Left,
        TableAlignment.Center => W.TableRowAlignmentValues.Center,
        TableAlignment.Right => W.TableRowAlignmentValues.Right,
        _ => throw new ArgumentOutOfRangeException(nameof(value), value, "The table alignment is not a defined value."),
    };

    private static TableLayout ToTableLayout(string value) => value switch
    {
        "autofit" => TableLayout.AutoFit,
        "fixed" => TableLayout.Fixed,
        _ => throw new InvalidOfficeFileException(
            $"The table stores the layout '{value}' that this document cannot read."),
    };

    private static W.TableLayoutValues ToOoxmlLayout(TableLayout value) => value switch
    {
        TableLayout.AutoFit => W.TableLayoutValues.Autofit,
        TableLayout.Fixed => W.TableLayoutValues.Fixed,
        _ => throw new ArgumentOutOfRangeException(nameof(value), value, "The table layout is not a defined value."),
    };

    /// <summary>Builds the <c>w:tbl</c> element for a new <paramref name="rows"/> by
    /// <paramref name="columns"/> table.</summary>
    internal static W.Table Create(int rows, int columns)
    {
        W.TableGrid grid = new();
        string width = DefaultColumnWidthTwips.ToString(CultureInfo.InvariantCulture);
        for (int column = 0; column < columns; column++)
        {
            grid.AppendChild(new W.GridColumn { Width = width });
        }

        // CT_Tbl requires w:tblPr and w:tblGrid before any w:tr. An empty w:tblPr keeps the created
        // table free of the properties (width, borders, shading) this version does not model yet.
        W.Table table = new(new W.TableProperties(), grid);
        for (int row = 0; row < rows; row++)
        {
            table.AppendChild(CreateRow(columns));
        }

        return table;
    }

    private static W.TableRow CreateRow(int columns)
    {
        W.TableRow row = new();
        for (int column = 0; column < columns; column++)
        {
            row.AppendChild(CreateCell());
        }

        return row;
    }

    private static W.TableCell CreateCell() => new(new W.Paragraph());
}
