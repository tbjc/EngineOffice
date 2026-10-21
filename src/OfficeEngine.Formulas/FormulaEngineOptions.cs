namespace OfficeEngine.Formulas;

/// <summary>The settings of the formula engine.</summary>
public sealed class FormulaEngineOptions
{
    /// <summary>
    /// The maximum length of a formula, in characters. The default is 8192, the same limit that
    /// Excel applies.
    /// </summary>
    public int MaximumFormulaLength { get; init; } = 8192;
}
