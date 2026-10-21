namespace OfficeEngine.Formulas;

/// <summary>A range of cells that a formula refers to.</summary>
internal sealed class FormulaRange
{
    internal FormulaRange(string? sheetName, int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        SheetName = sheetName;
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        LastRow = lastRow;
        LastColumn = lastColumn;
    }

    internal string? SheetName { get; }

    internal int FirstRow { get; }

    internal int FirstColumn { get; }

    internal int LastRow { get; }

    internal int LastColumn { get; }
}
