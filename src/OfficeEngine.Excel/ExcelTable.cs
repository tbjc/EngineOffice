namespace OfficeEngine.Excel;

/// <summary>
/// A structured table of a worksheet: a range with a name, a header row and named columns.
/// </summary>
/// <remarks>
/// The style, the totals row and the formulas of a table are not modelled; they are preserved when
/// the workbook is saved untouched.
/// </remarks>
public sealed class ExcelTable
{
    internal ExcelTable(string name, string range, bool hasHeaderRow, IReadOnlyList<string> columnNames)
    {
        Name = name;
        Range = range;
        HasHeaderRow = hasHeaderRow;
        ColumnNames = columnNames;
    }

    /// <summary>Gets the name of the table, for example <c>Tabla1</c>.</summary>
    public string Name { get; }

    /// <summary>Gets the A1-style range of the table, for example <c>A1:C5</c>.</summary>
    public string Range { get; }

    /// <summary>Gets whether the table has a header row.</summary>
    public bool HasHeaderRow { get; }

    /// <summary>Gets the names of the columns, in order.</summary>
    public IReadOnlyList<string> ColumnNames { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Name} ({Range})";
}
