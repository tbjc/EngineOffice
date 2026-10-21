using OfficeEngine.Excel;
using OfficeEngine.Formulas;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Formulas;

/// <summary>
/// The formulas of the engine, compared with the values that an independent engine computes for the
/// same workbook. This is the differential evidence of the increment.
/// </summary>
public sealed class FormulaDifferentialTests
{
    private static readonly (string Formula, string? Expected)[] Cases =
    [
        ("=A1+A2", "5"),
        ("=A1*A2", "6"),
        ("=A1/A2", "0.666666666666667"),
        ("=A2-A1", "1"),
        ("=A1^2+A2^2", "13"),
        ("=SUM(A1:A2)", "5"),
        ("=AVERAGE(A1:A2)", "2.5"),
        ("=MIN(A1:A2)", "2"),
        ("=MAX(A1:A2)", "3"),
        ("=COUNT(A1:A2)", "2"),
        ("=COUNTA(A1:A3)", "3"),
        ("=PRODUCT(A1,A2)", "6"),
        ("=IF(A1>1,\"si\",\"no\")", "si"),
        ("=IF(A1>5,\"si\",\"no\")", "no"),
        ("=AND(TRUE,A1=2)", "TRUE"),
        ("=OR(FALSE,A1=2)", "TRUE"),
        ("=ABS(0-A2)", "3"),
        ("=ROUND(2.345,2)", "2.35"),
        ("=MOD(7,3)", "1"),
        ("=POWER(2,8)", "256"),
        ("=CONCATENATE(A3,\" mundo\")", "hola mundo"),
        ("=LEN(A3)", "4"),
        ("=UPPER(A3)", "HOLA"),
        ("=A3&\"!\"", "hola!"),
        ("=1=1", "TRUE"),
        ("=1<2", "TRUE"),
        ("=\"a\"=\"A\"", "TRUE"),
        ("=1/0", "#DIV/0!"),
        ("=\"a\"+1", "#VALUE!"),
        ("=IFERROR(1/0,\"fallo\")", "fallo"),
    ];

    [LibreOfficeFact]
    public void TheEngine_ComputesTheSameValuesAsAnIndependentEngine()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet["A1"].Value = 2d;
        sheet["A2"].Value = 3d;
        sheet["A3"].Value = "hola";
        sheet["B1"].Value = 5d;
        sheet["B2"].Value = 7d;

        FormulaEngine engine = new();
        ExcelFormulaContext context = new(workbook);
        for (int index = 0; index < Cases.Length; index++)
        {
            sheet[index + 1, 4].Formula = Cases[index].Formula;
        }

        using MemoryStream stream = new();
        workbook.Save(stream);
        string[] lines = LibreOfficeTestSupport.ToCsv(stream.ToArray());

        Assert.True(lines.Length >= Cases.Length, "The CSV has fewer rows than the cases.");

        for (int index = 0; index < Cases.Length; index++)
        {
            (string formula, string? expected) = Cases[index];
            string[] cells = lines[index].Split(',');
            string fromEngine = cells.Length > 3 ? Normalize(cells[3]) : string.Empty;

            object? actual = engine.Evaluate(formula, context);
            string rendered = Render(actual);

            Assert.Equal(expected, rendered);
            Assert.Equal(fromEngine, rendered);
        }
    }

    /// <summary>
    /// The engine writes the booleans and the error values as Excel does, in English. The exporter
    /// writes them in the language of its interface, so those words are translated before comparing:
    /// the difference is a language, not a meaning.
    /// </summary>
    private static string Normalize(string value) => value switch
    {
        "VERDADERO" => "TRUE",
        "FALSO" => "FALSE",
        "#¡DIV/0!" => "#DIV/0!",
        "#¡VALOR!" => "#VALUE!",
        "#¡REF!" => "#REF!",
        "#¡NOMBRE?" => "#NAME?",
        "#¡N/D" => "#N/A",
        "#¡NUM!" => "#NUM!",
        "#¡NULO!" => "#NULL!",
        _ => value,
    };

    private static string Render(object? value) => value switch
    {
        null => string.Empty,
        double number => number.ToString("G15", System.Globalization.CultureInfo.InvariantCulture),
        bool flag => flag ? "TRUE" : "FALSE",
        string text => text,
        FormulaError error => error.Code,
        _ => value.ToString() ?? string.Empty,
    };
}
