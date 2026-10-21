using OfficeEngine.Formulas;

namespace OfficeEngine.Tests.Formulas;

/// <summary>The translation of the Spanish names of the functions into the canonical ones.</summary>
public sealed class FormulaNameTests
{
    [Theory]
    [InlineData("=SUMA(A1:A2)", "=SUM(A1:A2)")]
    [InlineData("=SI(A1>1,\"si\",\"no\")", "=IF(A1>1,\"si\",\"no\")")]
    [InlineData("=REDONDEAR.MAS(A1,2)", "=ROUNDUP(A1,2)")]
    [InlineData("=A1&B1", "=A1&B1")]
    [InlineData("=Hoja1!A1", "=Hoja1!A1")]
    [InlineData("='Mi Hoja'!A1", "='Mi Hoja'!A1")]
    public void Canonicalize_TranslatesTheNameOfAFunction(string formula, string expected)
    {
        Assert.Equal(expected, FormulaNames.Canonicalize(formula));
    }

    [Fact]
    public void Canonicalize_DoesNotTranslateTheTextOfALiteral()
    {
        Assert.Equal("=CONCATENATE(\"SUMA(\",A1)", FormulaNames.Canonicalize("=CONCATENAR(\"SUMA(\",A1)"));
    }
}
