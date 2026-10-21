namespace OfficeEngine.Excel;

/// <summary>The kinds of chart that a worksheet can hold.</summary>
public enum ExcelChartType
{
    /// <summary>A column chart: vertical bars grouped by category.</summary>
    Column = 0,

    /// <summary>A bar chart: horizontal bars grouped by category.</summary>
    Bar = 1,

    /// <summary>A line chart.</summary>
    Line = 2,

    /// <summary>A pie chart.</summary>
    Pie = 3,
}
