namespace OfficeEngine.Excel;

/// <summary>
/// A live, rectangular view over a range of cells of a <see cref="Worksheet"/>.
/// </summary>
/// <remarks>
/// <para>
/// A range is created with <see cref="Worksheet.GetRange(string)"/> or
/// <see cref="Worksheet.GetRange(int, int, int, int)"/> and includes both corners.
/// <see cref="Reference"/> uses the A1 form <c>B2:D5</c>.
/// </para>
/// <para>
/// The range is a live view: <see cref="Cells"/>, <see cref="GetValues"/> and
/// <see cref="SetValues(object?[,])"/> read and write the workbook each time they are called. Reading
/// a cell that does not exist yields <see langword="null"/> and does not create it; writing a value
/// creates the cell on demand.
/// </para>
/// <para>
/// Merged cells, styles and formulas are not modelled by the range. Each cell is addressed
/// independently, so a range over a merged area reports the values stored in the underlying cells.
/// <see cref="GetValues"/> returns a <see cref="DateTime"/> for cells whose number format is a
/// date/time format; <see cref="SetValues"/> does not write <see cref="DateTime"/> values because
/// each cell needs an explicit format. Use <see cref="Cell.SetDateTime"/> for those cells.
/// </para>
/// </remarks>
public sealed class CellRange
{
    private readonly ExcelWorkbook _workbook;
    private readonly Worksheet _worksheet;

    internal CellRange(
        ExcelWorkbook workbook,
        Worksheet worksheet,
        int firstRow,
        int firstColumn,
        int lastRow,
        int lastColumn)
    {
        _workbook = workbook;
        _worksheet = worksheet;
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        LastRow = lastRow;
        LastColumn = lastColumn;
        Reference = $"{CellReference.ToReference(firstColumn, firstRow)}:" +
            $"{CellReference.ToReference(lastColumn, lastRow)}";
    }

    /// <summary>Gets the A1-style reference of the range, for example <c>B2:D5</c>.</summary>
    public string Reference { get; }

    /// <summary>Gets the 1-based row index of the top-left corner.</summary>
    public int FirstRow { get; }

    /// <summary>Gets the 1-based column index of the top-left corner.</summary>
    public int FirstColumn { get; }

    /// <summary>Gets the 1-based row index of the bottom-right corner.</summary>
    public int LastRow { get; }

    /// <summary>Gets the 1-based column index of the bottom-right corner.</summary>
    public int LastColumn { get; }

    /// <summary>Gets the number of rows in the range.</summary>
    public int RowCount => LastRow - FirstRow + 1;

    /// <summary>Gets the number of columns in the range.</summary>
    public int ColumnCount => LastColumn - FirstColumn + 1;

    /// <summary>Gets the cells of the range in row-major order, starting at the top-left corner.</summary>
    /// <remarks>
    /// The sequence is lazy and reads the live workbook on every enumeration. Reading it does not
    /// create missing cells.
    /// </remarks>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public IEnumerable<Cell> Cells
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return EnumerateCells();
        }
    }

    /// <summary>Reads the values of the range as a two-dimensional array.</summary>
    /// <returns>
    /// An array with <see cref="RowCount"/> rows and <see cref="ColumnCount"/> columns, whose indices
    /// are 0-based offsets from the top-left corner. A cell that does not exist yields
    /// <see langword="null"/>.
    /// </returns>
    /// <remarks>
    /// The array is allocated in full; a very large range can exhaust memory. Read it in smaller
    /// ranges when its size is not bounded by the caller.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">A stored cell is malformed and cannot be read.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public object?[,] GetValues()
    {
        _workbook.ThrowIfDisposed();
        // shortcut: reads cell by cell, so cost grows superlinearly with the range;
        // upgrade when range throughput is measured as a bottleneck (Phase 5).
        object?[,] values = new object?[RowCount, ColumnCount];
        for (int row = 0; row < RowCount; row++)
        {
            for (int column = 0; column < ColumnCount; column++)
            {
                values[row, column] = _worksheet.GetCell(FirstRow + row, FirstColumn + column).Value;
            }
        }

        return values;
    }

    /// <summary>Writes the values of the range from a two-dimensional array.</summary>
    /// <param name="values">
    /// An array with exactly <see cref="RowCount"/> rows and <see cref="ColumnCount"/> columns, whose
    /// indices are 0-based offsets from the top-left corner. <see langword="null"/> clears an existing
    /// cell and does not create a missing one.
    /// </param>
    /// <remarks>
    /// Every value is validated before the first cell is written, so an unsupported value leaves the
    /// range unchanged. A <see cref="DateTime"/> is not accepted here because each cell needs an
    /// explicit date/time number format; use <see cref="Cell.SetDateTime"/> for those cells.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="values"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="values"/> does not have the required dimensions.
    /// </exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// A value has a type that a cell cannot store, including <see cref="DateTime"/>.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public void SetValues(object?[,] values)
    {
        ArgumentNullException.ThrowIfNull(values);
        _workbook.ThrowIfDisposed();

        if (values.GetLength(0) != RowCount || values.GetLength(1) != ColumnCount)
        {
            throw new ArgumentException(
                $"The array must have {RowCount} row(s) and {ColumnCount} column(s).",
                nameof(values));
        }

        for (int row = 0; row < RowCount; row++)
        {
            for (int column = 0; column < ColumnCount; column++)
            {
                Cell.EnsureRangeValue(values[row, column]);
            }
        }

        // shortcut: writes cell by cell, so cost grows superlinearly with the range;
        // upgrade when range throughput is measured as a bottleneck (Phase 5).
        for (int row = 0; row < RowCount; row++)
        {
            for (int column = 0; column < ColumnCount; column++)
            {
                Cell cell = _worksheet.GetCell(FirstRow + row, FirstColumn + column);
                object? value = values[row, column];
                if (value is null)
                {
                    cell.ClearValue();
                }
                else
                {
                    cell.Value = value;
                }
            }
        }
    }

    /// <inheritdoc />
    public override string ToString() => Reference;

    private IEnumerable<Cell> EnumerateCells()
    {
        for (int row = FirstRow; row <= LastRow; row++)
        {
            for (int column = FirstColumn; column <= LastColumn; column++)
            {
                yield return _worksheet.GetCell(row, column);
            }
        }
    }
}
