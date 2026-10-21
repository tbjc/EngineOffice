using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over a single cell (<c>w:tc</c>) of a <see cref="TableRow"/>.
/// </summary>
/// <remarks>
/// <para>
/// A cell holds block-level content: one or more paragraphs (and, in a document another tool wrote, a
/// nested table). This version exposes the cell's paragraphs and their text, together with the cell's
/// own formatting: preferred width, vertical alignment, shading and the four outer borders. A nested
/// table's content, cell merge (<c>w:gridSpan</c>/<c>w:vMerge</c>) and cell margins are preserved but
/// not reported here.
/// </para>
/// <para>
/// This is a lightweight wrapper: it does not own the underlying content and must not be used after
/// the <see cref="WordDocument"/> that produced it has been disposed.
/// </para>
/// </remarks>
public sealed class TableCell
{
    // Child order required by CT_TcPr. The names are the XML local names and the sequence is the one
    // the Open XML SDK generates for a TableCellProperties element, verified with a reflection probe
    // over DocumentFormat.OpenXml 3.5.1. Properties are inserted following this table instead of
    // being appended, so a cell that already carries markup this version does not model (a merge, a
    // margin) stays valid.
    private static readonly string[] CellPropertyOrder =
    [
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
        "textDirection", "tcFitText", "vAlign", "hideMark",
    ];

    private readonly W.TableCell _element;

    internal TableCell(W.TableCell element) => _element = element;

    private W.TableCellProperties? Properties => _element.GetFirstChild<W.TableCellProperties>();

    private W.TableCellBorders? Borders => Properties?.GetFirstChild<W.TableCellBorders>();

    /// <summary>Gets a live snapshot of the cell's paragraphs, in document order.</summary>
    /// <remarks>
    /// Only paragraphs that are direct children of the cell are returned. Content of a nested table
    /// that a foreign document put in this cell (which this version does not model) is not returned,
    /// although it is preserved when the document is saved.
    /// </remarks>
    public IReadOnlyList<Paragraph> Paragraphs
        => _element.Elements<W.Paragraph>().Select(element => new Paragraph(element)).ToArray();

    /// <summary>
    /// Gets or sets the visible text of the cell.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The getter concatenates the text of the cell's paragraphs (see <see cref="Paragraphs"/>),
    /// separated by line-feed characters, one per paragraph. Text inside a nested table is not
    /// included.
    /// </para>
    /// <para>
    /// The setter replaces the cell's paragraphs with a single paragraph carrying the new text. The
    /// cell properties (<c>w:tcPr</c>), such as width or shading, and block content this version does
    /// not model — a nested table, for example — are preserved. Only the paragraphs that held the old
    /// text are removed.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="value"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store. Add tabs and breaks with <see cref="Run.AddTab"/> and <see cref="Run.AddBreak"/>.
    /// </exception>
    public string Text
    {
        get => string.Join("\n", _element.Elements<W.Paragraph>().Select(paragraph => new Paragraph(paragraph).Text));

        set
        {
            ArgumentNullException.ThrowIfNull(value);
            PlainText.EnsureValid(value, nameof(value));

            // Remove only the paragraphs: everything else in the cell (w:tcPr, a nested table) is kept
            // so that assigning text never discards content this version does not model.
            foreach (W.Paragraph paragraph in _element.Elements<W.Paragraph>().ToArray())
            {
                paragraph.Remove();
            }

            _element.AppendChild(CreateParagraph(value));
        }
    }

    /// <summary>Adds a new paragraph with the given text at the end of the cell.</summary>
    /// <param name="text">The text of the paragraph. May be empty.</param>
    /// <returns>A live view over the new paragraph.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="text"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store. Add tabs and breaks with <see cref="Run.AddTab"/> and <see cref="Run.AddBreak"/>.
    /// </exception>
    public Paragraph AddParagraph(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        PlainText.EnsureValid(text, nameof(text));

        W.Paragraph paragraph = CreateParagraph(text);
        _element.AppendChild(paragraph);
        return new Paragraph(paragraph);
    }

    /// <summary>
    /// Gets or sets the preferred width of the cell.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the cell does not state a preferred width, which is exactly what
    /// Office Open XML expresses by omitting <c>w:tcW</c>; the cell then follows the width the table
    /// grid gives its column.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">A stored width cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The cell stores a width that Office Open XML defines but this version does not model.
    /// </exception>
    public TableWidth? Width
    {
        get => TableWidth.Read(Properties?.GetFirstChild<W.TableCellWidth>(), "cell width");
        set
        {
            if (value is not { } width)
            {
                RemoveProperty<W.TableCellWidth>();
                return;
            }

            width.ApplyTo(EnsureProperty<W.TableCellWidth>());
        }
    }

    /// <summary>
    /// Gets or sets the number of grid columns the cell covers (<c>w:gridSpan</c>).
    /// </summary>
    /// <remarks>
    /// <para>
    /// A cell that states no span covers one column, which is exactly what Office Open XML expresses by
    /// omitting <c>w:gridSpan</c>; a cell merged with the ones beside it covers as many columns as its
    /// span says, and the row holds one cell fewer per merged column.
    /// </para>
    /// <para>
    /// Assigning <c>1</c> removes <c>w:gridSpan</c>, because one column and no span are the same thing.
    /// <see cref="TableRow.MergeCells"/> and <see cref="TableRow.UnmergeCell"/> set the span and also
    /// move the content of, or create, the cells it covers.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is less than one.</exception>
    /// <exception cref="InvalidOfficeFileException">
    /// A stored span is missing its <c>w:val</c> attribute or cannot be parsed.
    /// </exception>
    public int ColumnSpan
    {
        get
        {
            W.GridSpan? stored = Properties?.GetFirstChild<W.GridSpan>();
            return stored is null ? 1 : ReadColumnSpan(stored.Val?.InnerText);
        }

        set
        {
            if (value < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(value), value, "The column span must be at least one.");
            }

            if (value == 1)
            {
                RemoveProperty<W.GridSpan>();
                return;
            }

            EnsureProperty<W.GridSpan>().Val = value;
        }
    }

    /// <summary>
    /// Gets or sets the part the cell plays in a vertical merge (<c>w:vMerge</c>).
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the cell is not part of a vertical merge, which is exactly what
    /// Office Open XML expresses by omitting <c>w:vMerge</c>. A stored <c>w:vMerge</c> without
    /// <c>w:val</c> reads as <see cref="TableVerticalMerge.Continue"/>, the value Office Open XML
    /// assumes; this version always writes the value. Use <see cref="Table.MergeCellsVertically"/> and
    /// <see cref="Table.UnmergeCellsVertically"/> to merge and split whole runs of rows.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is not a defined state.</exception>
    /// <exception cref="InvalidOfficeFileException">A stored state cannot be read.</exception>
    public TableVerticalMerge? VerticalMerge
    {
        get
        {
            W.VerticalMerge? stored = Properties?.GetFirstChild<W.VerticalMerge>();
            if (stored is null)
            {
                return null;
            }

            // ST_Merge: an omitted w:val means the cell continues the merge that starts above it.
            return ToTableVerticalMerge(stored.Val?.InnerText ?? "continue");
        }

        set
        {
            if (value is not { } merge)
            {
                RemoveProperty<W.VerticalMerge>();
                return;
            }

            // Map before touching w:tcPr so that a rejected value leaves no partial markup.
            W.MergedCellValues ooxml = ToOoxmlVerticalMerge(merge);
            EnsureProperty<W.VerticalMerge>().Val = ooxml;
        }
    }

    /// <summary>
    /// Gets or sets the vertical alignment of the cell's content.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the cell inherits its vertical alignment from the table and the
    /// defaults, which is exactly what Office Open XML expresses by omitting <c>w:vAlign</c>.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is not a defined alignment.</exception>
    /// <exception cref="InvalidOfficeFileException">A stored alignment cannot be read.</exception>
    public TableVerticalAlignment? VerticalAlignment
    {
        get
        {
            string? stored = Properties?.GetFirstChild<W.TableCellVerticalAlignment>()?.Val?.InnerText;
            return stored is null ? null : ToTableVerticalAlignment(stored);
        }

        set
        {
            if (value is not { } alignment)
            {
                RemoveProperty<W.TableCellVerticalAlignment>();
                return;
            }

            // Map before touching w:tcPr so that a rejected value leaves no partial markup.
            W.TableVerticalAlignmentValues ooxml = ToOoxmlVerticalAlignment(alignment);
            EnsureProperty<W.TableCellVerticalAlignment>().Val = ooxml;
        }
    }

    /// <summary>
    /// Gets or sets the shading colour of the cell.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see langword="null"/> means the cell states no shading, which is exactly what Office Open XML
    /// expresses by omitting <c>w:shd</c>. A cell that carries only a theme fill, which this version
    /// does not model, also reads as <see langword="null"/>.
    /// </para>
    /// <para>
    /// Assigning a colour writes a solid fill (the <c>clear</c> pattern). Assigning
    /// <see langword="null"/> removes the shading this cell states.
    /// </para>
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">A stored shading cannot be read.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The cell stores a shading pattern that Office Open XML defines but this version does not model.
    /// </exception>
    public RgbColor? ShadingColor
    {
        get => TableShading.Read(Properties?.GetFirstChild<W.Shading>());
        set
        {
            if (value is not { } color)
            {
                RemoveProperty<W.Shading>();
                return;
            }

            TableShading.Write(EnsureProperty<W.Shading>(), color);
        }
    }

    /// <summary>Gets or sets the top border of the cell.</summary>
    /// <remarks>
    /// <see langword="null"/> means the cell does not state this edge, so it follows the table's
    /// borders and the defaults; <c>w:tcBorders</c> is created when the first edge is set and removed
    /// again when the last one is cleared. The edges a cell does not own — the inside borders of the
    /// table and the diagonal edges — are never rewritten.
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

    /// <summary>Gets or sets the left border of the cell.</summary>
    /// <remarks>
    /// See <see cref="BorderTop"/> for the meaning of <see langword="null"/> and for the failures
    /// reading a stored border can raise.
    /// </remarks>
    public TableBorder? BorderLeft
    {
        get => BorderEdges.Read<W.LeftBorder>(Borders);
        set => SetBorder<W.LeftBorder>(value);
    }

    /// <summary>Gets or sets the bottom border of the cell.</summary>
    /// <remarks>
    /// See <see cref="BorderTop"/> for the meaning of <see langword="null"/> and for the failures
    /// reading a stored border can raise.
    /// </remarks>
    public TableBorder? BorderBottom
    {
        get => BorderEdges.Read<W.BottomBorder>(Borders);
        set => SetBorder<W.BottomBorder>(value);
    }

    /// <summary>Gets or sets the right border of the cell.</summary>
    /// <remarks>
    /// See <see cref="BorderTop"/> for the meaning of <see langword="null"/> and for the failures
    /// reading a stored border can raise.
    /// </remarks>
    public TableBorder? BorderRight
    {
        get => BorderEdges.Read<W.RightBorder>(Borders);
        set => SetBorder<W.RightBorder>(value);
    }

    /// <inheritdoc />
    public override string ToString() => $"table cell (paragraphs: {Paragraphs.Count})";

    private W.TableCellProperties EnsureProperties()
    {
        W.TableCellProperties? properties = _element.GetFirstChild<W.TableCellProperties>();
        if (properties is null)
        {
            // CT_Tc requires w:tcPr, when present, to be the first child of the cell.
            properties = new W.TableCellProperties();
            _element.InsertAt(properties, 0);
        }

        return properties;
    }

    private TChild EnsureProperty<TChild>()
        where TChild : OpenXmlElement, new()
    {
        W.TableCellProperties properties = EnsureProperties();
        TChild? existing = properties.GetFirstChild<TChild>();
        if (existing is not null)
        {
            return existing;
        }

        TChild created = new();
        SchemaOrder.Insert(properties, created, CellPropertyOrder);
        return created;
    }

    private void RemoveProperty<TChild>()
        where TChild : OpenXmlElement
    {
        W.TableCellProperties? properties = Properties;
        if (properties is null)
        {
            return;
        }

        properties.RemoveAllChildren<TChild>();
        PrunePropertiesIfEmpty(properties);
    }

    private void SetBorder<TEdge>(TableBorder? value)
        where TEdge : W.BorderType, new()
    {
        if (value is not { } border)
        {
            W.TableCellProperties? properties = Properties;
            W.TableCellBorders? borders = properties?.GetFirstChild<W.TableCellBorders>();
            if (properties is null || borders is null)
            {
                return;
            }

            BorderEdges.Write<TEdge>(borders, null);
            if (!borders.HasChildren)
            {
                borders.Remove();
                PrunePropertiesIfEmpty(properties);
            }

            return;
        }

        BorderEdges.Write<TEdge>(EnsureProperty<W.TableCellBorders>(), border);
    }

    private static void PrunePropertiesIfEmpty(W.TableCellProperties properties)
    {
        if (!properties.HasChildren && !properties.GetAttributes().Any())
        {
            properties.Remove();
        }
    }

    private static int ReadColumnSpan(string? value)
        => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span) && span >= 1
            ? span
            : throw new InvalidOfficeFileException(
                $"The cell stores the grid span '{value}' that this document cannot read.");

    private static TableVerticalAlignment ToTableVerticalAlignment(string value) => value switch
    {
        "top" => TableVerticalAlignment.Top,
        "center" => TableVerticalAlignment.Center,
        "bottom" => TableVerticalAlignment.Bottom,
        _ => throw new InvalidOfficeFileException(
            $"The cell stores the vertical alignment '{value}' that this document cannot read."),
    };

    private static W.TableVerticalAlignmentValues ToOoxmlVerticalAlignment(TableVerticalAlignment value) => value switch
    {
        TableVerticalAlignment.Top => W.TableVerticalAlignmentValues.Top,
        TableVerticalAlignment.Center => W.TableVerticalAlignmentValues.Center,
        TableVerticalAlignment.Bottom => W.TableVerticalAlignmentValues.Bottom,
        _ => throw new ArgumentOutOfRangeException(
            nameof(value),
            value,
            "The cell vertical alignment is not a defined value."),
    };

    private static TableVerticalMerge ToTableVerticalMerge(string value) => value switch
    {
        "continue" => TableVerticalMerge.Continue,
        "restart" => TableVerticalMerge.Restart,
        _ => throw new InvalidOfficeFileException(
            $"The cell stores the vertical merge '{value}' that this document cannot read."),
    };

    private static W.MergedCellValues ToOoxmlVerticalMerge(TableVerticalMerge value) => value switch
    {
        TableVerticalMerge.Continue => W.MergedCellValues.Continue,
        TableVerticalMerge.Restart => W.MergedCellValues.Restart,
        _ => throw new ArgumentOutOfRangeException(
            nameof(value),
            value,
            "The cell vertical merge is not a defined value."),
    };

    private static W.Paragraph CreateParagraph(string text)
        => text.Length == 0
            ? new W.Paragraph()
            : new W.Paragraph(new W.Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
}
