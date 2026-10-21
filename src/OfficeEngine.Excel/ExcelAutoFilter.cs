namespace OfficeEngine.Excel;

/// <summary>The auto filter of a worksheet: the range of cells that the filter buttons cover.</summary>
/// <remarks>
/// The filter criteria themselves (which rows are hidden) are not modelled; they are preserved when
/// the worksheet is saved untouched and dropped when the filter is written again.
/// </remarks>
public sealed class ExcelAutoFilter
{
    /// <summary>Initializes a new instance of the <see cref="ExcelAutoFilter"/> class.</summary>
    /// <param name="rangeReference">
    /// The A1-style range the filter covers, including the header row, for example <c>A1:D10</c>.
    /// </param>
    /// <exception cref="ArgumentException"><paramref name="rangeReference"/> is not a valid A1-style reference.</exception>
    public ExcelAutoFilter(string rangeReference)
    {
        Range = CellReference.NormalizeRangeOrCell(rangeReference, nameof(rangeReference));
    }

    /// <summary>Gets the A1-style range the filter covers, for example <c>A1:D10</c>.</summary>
    public string Range { get; }

    /// <inheritdoc />
    public override string ToString() => Range;
}
