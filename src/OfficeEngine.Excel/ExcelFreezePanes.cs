namespace OfficeEngine.Excel;

/// <summary>The frozen part of a worksheet: how many rows and columns stay visible while scrolling.</summary>
public sealed class ExcelFreezePanes
{
    /// <summary>Initializes a new instance of the <see cref="ExcelFreezePanes"/> class.</summary>
    /// <param name="rows">The number of frozen rows, at the top of the sheet.</param>
    /// <param name="columns">The number of frozen columns, at the left of the sheet.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A count is negative or greater than the number of rows or columns of a worksheet.
    /// </exception>
    /// <exception cref="ArgumentException">Both counts are zero, which would freeze nothing.</exception>
    public ExcelFreezePanes(int rows, int columns)
    {
        if (rows < 0 || rows > CellReference.MaxRowIndex)
        {
            throw new ArgumentOutOfRangeException(nameof(rows), rows, "The number of frozen rows is out of range.");
        }

        if (columns < 0 || columns > CellReference.MaxColumnIndex)
        {
            throw new ArgumentOutOfRangeException(
                nameof(columns), columns, "The number of frozen columns is out of range.");
        }

        if (rows == 0 && columns == 0)
        {
            throw new ArgumentException("At least one row or column must be frozen.", nameof(rows));
        }

        Rows = rows;
        Columns = columns;
    }

    /// <summary>Gets the number of frozen rows.</summary>
    public int Rows { get; }

    /// <summary>Gets the number of frozen columns.</summary>
    public int Columns { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Columns} columnas y {Rows} filas congeladas";

    internal string TopLeftCell => CellReference.ToReference(Columns + 1, Rows + 1);

    internal string ActivePane => (Rows > 0, Columns > 0) switch
    {
        (true, true) => "bottomRight",
        (true, false) => "bottomLeft",
        _ => "topRight",
    };
}
