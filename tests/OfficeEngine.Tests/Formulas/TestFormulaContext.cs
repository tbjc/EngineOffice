using OfficeEngine.Formulas;

namespace OfficeEngine.Tests.Formulas;

/// <summary>A context with a fixed set of cells, for the tests of the engine.</summary>
internal sealed class TestFormulaContext : IFormulaContext
{
    private readonly Dictionary<string, object?> _cells;

    internal TestFormulaContext(string? sheetName = null)
    {
        SheetName = sheetName;
        _cells = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["A1"] = 2d,
            ["A2"] = 3d,
            ["A3"] = "hola",
            ["A4"] = true,
            ["A5"] = 10d,
            ["A6"] = 20d,
            ["B1"] = 5d,
            ["B2"] = 7d,
        };
    }

    public string? SheetName { get; }

    public bool TryGetCellValue(string? sheetName, int row, int column, out object? value)
    {
        value = null;
        string reference = $"{ColumnName(column)}{row}";
        return _cells.TryGetValue(reference, out value);
    }

    internal static string ColumnName(int column)
    {
        string name = string.Empty;
        int value = column;
        while (value > 0)
        {
            int remainder = (value - 1) % 26;
            name = (char)('A' + remainder) + name;
            value = (value - 1) / 26;
        }

        return name;
    }
}
