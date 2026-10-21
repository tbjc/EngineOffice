namespace OfficeEngine.Excel;

/// <summary>A chart anchored in a worksheet, with a single series.</summary>
/// <remarks>
/// The chart points at the ranges of the sheet that hold the categories and the values, and it
/// carries a copy of them (the caché) so that it can be shown without recalculating.
/// </remarks>
public sealed class ExcelChart
{
    internal ExcelChart(
        string name,
        string reference,
        ExcelChartType type,
        string categoriesRange,
        string valuesRange,
        string? title,
        string? seriesName)
    {
        Name = name;
        Reference = reference;
        Type = type;
        CategoriesRange = categoriesRange;
        ValuesRange = valuesRange;
        Title = title;
        SeriesName = seriesName;
    }

    /// <summary>Gets the name of the chart, for example <c>Gráfico 1</c>.</summary>
    public string Name { get; }

    /// <summary>Gets the A1-style reference of the cell the chart is anchored to.</summary>
    public string Reference { get; }

    /// <summary>Gets the kind of chart.</summary>
    public ExcelChartType Type { get; }

    /// <summary>Gets the A1-style range that holds the categories, for example <c>A2:A5</c>.</summary>
    public string CategoriesRange { get; }

    /// <summary>Gets the A1-style range that holds the values, for example <c>B2:B5</c>.</summary>
    public string ValuesRange { get; }

    /// <summary>Gets the title of the chart, when it has one.</summary>
    public string? Title { get; }

    /// <summary>Gets the name of the series, when it has one.</summary>
    public string? SeriesName { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Name} ({Type})";
}
