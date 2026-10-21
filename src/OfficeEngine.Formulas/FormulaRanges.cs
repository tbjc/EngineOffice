namespace OfficeEngine.Formulas;

/// <summary>Reads the values of the cells of a range.</summary>
internal static class FormulaRanges
{
    /// <summary>Enumerates the values of the cells of a range, in row-major order.</summary>
    /// <param name="range">The range to read.</param>
    /// <param name="context">The context that supplies the values.</param>
    /// <returns>The value of every cell, which is <see langword="null"/> for the empty ones.</returns>
    internal static IEnumerable<object?> Values(FormulaRange range, IFormulaContext context)
    {
        for (int row = range.FirstRow; row <= range.LastRow; row++)
        {
            for (int column = range.FirstColumn; column <= range.LastColumn; column++)
            {
                context.TryGetCellValue(range.SheetName, row, column, out object? value);
                yield return value;
            }
        }
    }
}
