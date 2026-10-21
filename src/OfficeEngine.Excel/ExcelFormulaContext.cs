using OfficeEngine.Formulas;

namespace OfficeEngine.Excel;

/// <summary>
/// Supplies the values of a workbook to the formula engine.
/// </summary>
/// <remarks>
/// The context reads the cells of the workbook as they are: a formula is evaluated against the
/// values that the workbook has, not against the formulas that other cells contain.
/// </remarks>
public sealed class ExcelFormulaContext : IFormulaContext
{
    private readonly ExcelWorkbook _workbook;

    /// <summary>Initializes a new instance of the <see cref="ExcelFormulaContext"/> class.</summary>
    /// <param name="workbook">The workbook that supplies the values.</param>
    /// <param name="sheetName">
    /// The sheet the formulas belong to, used by the references without a sheet. The first sheet is
    /// used when it is not given.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">The sheet does not exist in the workbook.</exception>
    public ExcelFormulaContext(ExcelWorkbook workbook, string? sheetName = null)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        _workbook = workbook;
        SheetName = sheetName ?? workbook.Worksheets[0].Name;
    }

    /// <inheritdoc />
    public string? SheetName { get; }

    /// <inheritdoc />
    public bool TryGetCellValue(string? sheetName, int row, int column, out object? value)
    {
        value = null;
        string name = sheetName ?? SheetName ?? string.Empty;
        if (!_workbook.Worksheets.Any(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        value = _workbook.Worksheets[name].GetCell(row, column).Value;
        return true;
    }
}
