namespace OfficeEngine.Excel;

/// <summary>The print options of a sheet: what is printed and how it is placed on the page.</summary>
public sealed class ExcelPrintOptions
{
    /// <summary>Initializes a new instance of the <see cref="ExcelPrintOptions"/> class.</summary>
    /// <param name="horizontalCentered">Whether the content is centered horizontally.</param>
    /// <param name="verticalCentered">Whether the content is centered vertically.</param>
    /// <param name="gridLines">Whether the grid lines are printed.</param>
    /// <param name="headings">Whether the row and column headings are printed.</param>
    public ExcelPrintOptions(
        bool horizontalCentered = false,
        bool verticalCentered = false,
        bool gridLines = false,
        bool headings = false)
    {
        HorizontalCentered = horizontalCentered;
        VerticalCentered = verticalCentered;
        GridLines = gridLines;
        Headings = headings;
    }

    /// <summary>Gets whether the content is centered horizontally.</summary>
    public bool HorizontalCentered { get; }

    /// <summary>Gets whether the content is centered vertically.</summary>
    public bool VerticalCentered { get; }

    /// <summary>Gets whether the grid lines are printed.</summary>
    public bool GridLines { get; }

    /// <summary>Gets whether the row and column headings are printed.</summary>
    public bool Headings { get; }

    /// <inheritdoc />
    public override string ToString() => $"Centro: {HorizontalCentered}/{VerticalCentered}";
}
