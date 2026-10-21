using OfficeEngine.Formulas;

namespace OfficeEngine.Tests.Formulas;

/// <summary>The evaluation of formulas: the values, the operators and the errors.</summary>
public sealed class FormulaEngineTests
{
    [Theory]
    [InlineData("=1+2", 3d)]
    [InlineData("=10/4", 2.5d)]
    [InlineData("=2^10", 1024d)]
    [InlineData("=-3+1", -2d)]
    [InlineData("=50%", 0.5d)]
    [InlineData("=A1+A2", 5d)]
    [InlineData("=A1*B1", 10d)]
    [InlineData("=SUM(A1:A2,B1:B2)", 17d)]
    [InlineData("=AVERAGE(A1:A2)", 2.5d)]
    [InlineData("=MIN(A1:A2)", 2d)]
    [InlineData("=MAX(A1:A2)", 3d)]
    [InlineData("=COUNT(A1:A2)", 2d)]
    [InlineData("=COUNTA(A1:A3)", 3d)]
    [InlineData("=PRODUCT(A1,A2)", 6d)]
    [InlineData("=IF(A1>1,\"si\",\"no\")", "si")]
    [InlineData("=IF(A1>5,\"si\",\"no\")", "no")]
    [InlineData("=AND(TRUE,A1=2)", true)]
    [InlineData("=OR(FALSE,A1=2)", true)]
    [InlineData("=NOT(A1=2)", false)]
    [InlineData("=ABS(0-3)", 3d)]
    [InlineData("=INT(2.7)", 2d)]
    [InlineData("=SQRT(16)", 4d)]
    [InlineData("=ROUND(2.345,2)", 2.35d)]
    [InlineData("=ROUNDUP(2.341,2)", 2.35d)]
    [InlineData("=ROUNDDOWN(2.349,2)", 2.34d)]
    [InlineData("=MOD(7,3)", 1d)]
    [InlineData("=POWER(2,8)", 256d)]
    [InlineData("=CONCATENATE(A3,\" mundo\")", "hola mundo")]
    [InlineData("=LEN(A3)", 4d)]
    [InlineData("=LEFT(A3,2)", "ho")]
    [InlineData("=RIGHT(A3,2)", "la")]
    [InlineData("=MID(A3,2,2)", "ol")]
    [InlineData("=UPPER(A3)", "HOLA")]
    [InlineData("=LOWER(A3)", "hola")]
    [InlineData("=TRIM(\"  a  b  \")", "a b")]
    [InlineData("=ISBLANK(Z9)", true)]
    [InlineData("=ISNUMBER(A1)", true)]
    [InlineData("=ISTEXT(A3)", true)]
    [InlineData("=ISERROR(1/0)", true)]
    [InlineData("=IFERROR(1/0,\"fallo\")", "fallo")]
    [InlineData("=A3&\" mundo\"", "hola mundo")]
    [InlineData("=1=1", true)]
    [InlineData("=1<>2", true)]
    [InlineData("=1<2", true)]
    [InlineData("=\"a\"=\"A\"", true)]
    [InlineData("=TRUE", true)]
    [InlineData("=1+TRUE", 2d)]
    public void Evaluate_ComputesTheExpectedValue(string formula, object? expected)
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        object? actual = engine.Evaluate(formula, context);

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Evaluate_WithoutTheLeadingEquals_ComputesTheSameValue()
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        Assert.Equal(3d, engine.Evaluate("1+2", context));
    }

    [Theory]
    [InlineData("=1/0", "#DIV/0!")]
    [InlineData("=\"a\"+1", "#VALUE!")]
    [InlineData("=UNKNOWN()", "#NAME?")]
    [InlineData("=SQRT(-1)", "#NUM!")]
        public void Evaluate_ReportsTheErrorOfExcel(string formula, string expectedCode)
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        object? actual = engine.Evaluate(formula, context);

        FormulaError error = Assert.IsType<FormulaError>(actual);
        Assert.Equal(expectedCode, error.Code);
    }

    [Fact]
    public void Evaluate_DoesNotEvaluateTheBranchOfIfThatIsNotTaken()
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        object? actual = engine.Evaluate("=IF(A1>0,1,1/0)", context);

        Assert.Equal(1d, actual);
    }

    [Fact]
    public void Evaluate_ReadsTheCellsOfTheSheet()
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new("Datos");

        Assert.Equal(5d, engine.Evaluate("=A1+A2", context));
    }

    [Fact]
    public void Evaluate_WithAnEmptyFormula_Throws()
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        Assert.Throws<ArgumentException>(() => engine.Evaluate("  ", context));
    }

    [Fact]
    public void Evaluate_WithATooLongFormula_Throws()
    {
        FormulaEngine engine = new(new FormulaEngineOptions { MaximumFormulaLength = 10 });
        TestFormulaContext context = new();

        Assert.Throws<ArgumentException>(() => engine.Evaluate("=1+1+1+1+1+1+1", context));
    }

    [Fact]
    public void Evaluate_WithAnUnreadableFormula_Throws()
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        Assert.Throws<FormulaSyntaxException>(() => engine.Evaluate("=1+", context));
        Assert.Throws<FormulaSyntaxException>(() => engine.Evaluate("=SUM(", context));
    }

    [Fact]
    public void TheMatrixOfFunctions_IsPublished()
    {
        Assert.Contains("SUM", FormulaEngine.SupportedFunctions);
        Assert.Contains("IF", FormulaEngine.SupportedFunctions);
        Assert.Contains("CONCATENATE", FormulaEngine.SupportedFunctions);
        Assert.Equal(FormulaEngine.SupportedFunctions.Count, FormulaEngine.SupportedFunctions.Distinct().Count());
    }

    [Theory]
    [InlineData("=1e3", 1000d)]
    [InlineData("=1E-2", 0.01d)]
    [InlineData("=0.5+0.25", 0.75d)]
    [InlineData("=(1+2)*3", 9d)]
    [InlineData("=2^3^2", 512d)]
    [InlineData("=\"a\"&\"b\"&\"c\"", "abc")]
    [InlineData("=SUM(A1:A2)+A1", 7d)]
    public void Evaluate_ReadsTheOperatorsOfExcel(string formula, object? expected)
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        Assert.Equal(expected, engine.Evaluate(formula, context));
    }

    [Theory]
    [InlineData("=SUM()", "#VALUE!")]
    [InlineData("=NOT(TRUE,TRUE)", "#VALUE!")]
    [InlineData("=SQRT()", "#VALUE!")]
    public void Evaluate_ReportsTheWrongNumberOfArguments(string formula, string expectedCode)
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        object? actual = engine.Evaluate(formula, context);

        FormulaError error = Assert.IsType<FormulaError>(actual);
        Assert.Equal(expectedCode, error.Code);
    }

    [Fact]
    public void Evaluate_WithARangeInAnArithmeticOperation_Fails()
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        object? actual = engine.Evaluate("=A1:A2+1", context);

        FormulaError error = Assert.IsType<FormulaError>(actual);
        Assert.Equal("#VALUE!", error.Code);
    }

    [Fact]
    public void Evaluate_WithAnUnknownName_Fails()
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        object? actual = engine.Evaluate("=NOMBRE", context);

        FormulaError error = Assert.IsType<FormulaError>(actual);
        Assert.Equal("#NAME?", error.Code);
    }

    [Theory]
    [InlineData("=1+", "Unexpected")]
    [InlineData("=\"sin cerrar", "not closed")]
    [InlineData("=SUM(A1,", "Unexpected")]
    [InlineData("=1 @ 2", "Unexpected character")]
    public void Evaluate_WithAMalformedFormula_Throws(string formula, string expectedMessage)
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        FormulaSyntaxException error = Assert.Throws<FormulaSyntaxException>(() => engine.Evaluate(formula, context));
        Assert.Contains(expectedMessage, error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Evaluate_WithAFormulaThatUsesTheWholeText_Throws()
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        Assert.Throws<FormulaSyntaxException>(() => engine.Evaluate("=1 2", context));
    }

    [Fact]
    public void Evaluate_WithAnErrorValueInTheFormula_PropagatesIt()
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        object? actual = engine.Evaluate("=#DIV/0!+1", context);

        FormulaError error = Assert.IsType<FormulaError>(actual);
        Assert.Equal("#DIV/0!", error.Code);
    }

    [Theory]
    [InlineData("=SUMA(A1:A2)", "=SUM(A1:A2)")]
    [InlineData("=PRODUCTO(A1,A2)", "=PRODUCT(A1,A2)")]
    [InlineData("=PROMEDIO(A1:A2)", "=AVERAGE(A1:A2)")]
    [InlineData("=CONTAR(A1:A2)", "=COUNT(A1:A2)")]
    [InlineData("=CONTARA(A1:A3)", "=COUNTA(A1:A3)")]
    [InlineData("=Y(TRUE,A1=2)", "=AND(TRUE,A1=2)")]
    [InlineData("=O(FALSE,A1=2)", "=OR(FALSE,A1=2)")]
    [InlineData("=NO(A1=2)", "=NOT(A1=2)")]
    [InlineData("=SI(A1>1,\"si\",\"no\")", "=IF(A1>1,\"si\",\"no\")")]
    [InlineData("=SI.ERROR(1/0,\"fallo\")", "=IFERROR(1/0,\"fallo\")")]
    [InlineData("=ENTERO(2.7)", "=INT(2.7)")]
    [InlineData("=RAIZ(16)", "=SQRT(16)")]
    [InlineData("=REDONDEAR(2.345,2)", "=ROUND(2.345,2)")]
    [InlineData("=REDONDEAR.MAS(2.341,2)", "=ROUNDUP(2.341,2)")]
    [InlineData("=REDONDEAR.MENOS(2.349,2)", "=ROUNDDOWN(2.349,2)")]
    [InlineData("=RESIDUO(7,3)", "=MOD(7,3)")]
    [InlineData("=POTENCIA(2,8)", "=POWER(2,8)")]
    [InlineData("=CONCATENAR(A3,\" mundo\")", "=CONCATENATE(A3,\" mundo\")")]
    [InlineData("=LARGO(A3)", "=LEN(A3)")]
    [InlineData("=IZQUIERDA(A3,2)", "=LEFT(A3,2)")]
    [InlineData("=DERECHA(A3,2)", "=RIGHT(A3,2)")]
    [InlineData("=EXTRAE(A3,2,2)", "=MID(A3,2,2)")]
    [InlineData("=MAYUSC(A3)", "=UPPER(A3)")]
    [InlineData("=MINUSC(A3)", "=LOWER(A3)")]
    [InlineData("=ESPACIOS(\"  a  b  \")", "=TRIM(\"  a  b  \")")]
    [InlineData("=ESBLANCO(Z9)", "=ISBLANK(Z9)")]
    [InlineData("=ESNUMERO(A1)", "=ISNUMBER(A1)")]
    [InlineData("=ESTEXTO(A3)", "=ISTEXT(A3)")]
    [InlineData("=ESERROR(1/0)", "=ISERROR(1/0)")]
    public void Evaluate_WithTheSpanishNameOfAFunction_ComputesTheSameValue(string spanish, string english)
    {
        FormulaEngine engine = new();
        TestFormulaContext context = new();

        ArgumentNullException.ThrowIfNull(spanish);
        ArgumentNullException.ThrowIfNull(english);

        object? expected = engine.Evaluate(english, context);
        object? actual = engine.Evaluate(spanish, context);

        Assert.Equal(expected, actual);
        Assert.Contains(NameOf(spanish), FormulaEngine.SupportedFunctions);
        Assert.Contains(NameOf(english), FormulaEngine.SupportedFunctions);
    }

    private static string NameOf(string formula)
        => formula[1..formula.IndexOf('(', StringComparison.Ordinal)];
}
