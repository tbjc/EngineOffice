namespace OfficeEngine.Formulas;

/// <summary>
/// Supplies the values that a formula needs: the cells it refers to and the sheet it belongs to.
/// </summary>
/// <remarks>
/// The engine does not know about Word or Excel. A caller that has a workbook implements this
/// interface over it, and the engine asks for the values it needs.
/// </remarks>
public interface IFormulaContext
{
    /// <summary>Gets the name of the sheet the formula belongs to, used by the references without a sheet.</summary>
    string? SheetName { get; }

    /// <summary>Reads the value of a cell.</summary>
    /// <param name="sheetName">The sheet of the cell, or <see langword="null"/> for the current one.</param>
    /// <param name="row">The 1-based row of the cell.</param>
    /// <param name="column">The 1-based column of the cell.</param>
    /// <param name="value">The value of the cell, or <see langword="null"/> when it is empty.</param>
    /// <returns><see langword="true"/> when the cell exists; <see langword="false"/> when it is empty or does not exist.</returns>
    bool TryGetCellValue(string? sheetName, int row, int column, out object? value);
}
