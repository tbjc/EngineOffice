namespace OfficeEngine.Formulas;

/// <summary>
/// Evaluates formulas over the values of a context.
/// </summary>
/// <remarks>
/// <para>
/// The engine reads the formula, resolves the references it contains and applies the functions of the
/// published matrix. It never calculates anything that the formula does not ask for, and it does not
/// know about Word or Excel: the values come from an <see cref="IFormulaContext"/>.
/// </para>
/// <para>
/// The result is a <see cref="double"/>, a <see cref="string"/>, a <see cref="bool"/>,
/// a <see cref="FormulaError"/> or <see langword="null"/> for an empty value, the same set of values
/// that Excel uses.
/// </para>
/// </remarks>
public sealed class FormulaEngine
{
    private readonly FormulaEngineOptions _options;

    /// <summary>Initializes a new instance of the <see cref="FormulaEngine"/> class.</summary>
    /// <param name="options">The settings of the engine. The defaults are used when it is not given.</param>
    public FormulaEngine(FormulaEngineOptions? options = null)
        => _options = options ?? new FormulaEngineOptions();

    /// <summary>Gets the names of the functions this engine understands, in a stable order.</summary>
    public static IReadOnlyList<string> SupportedFunctions { get; } = FormulaFunctions.Names;

    /// <summary>Evaluates a formula.</summary>
    /// <param name="formula">
    /// The formula, with or without the leading <c>=</c>.
    /// </param>
    /// <param name="context">The context that supplies the values of the cells.</param>
    /// <returns>
    /// The value of the formula: a <see cref="double"/>, a <see cref="string"/>, a
    /// <see cref="bool"/>, a <see cref="FormulaError"/> or <see langword="null"/>.
    /// </returns>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="formula"/> or <paramref name="context"/> is <see langword="null"/>.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="formula"/> is empty or longer than the allowed length.
    /// </exception>
    /// <exception cref="FormulaSyntaxException">The formula cannot be read.</exception>
    public object? Evaluate(string formula, IFormulaContext context)
    {
        ArgumentNullException.ThrowIfNull(formula);
        ArgumentNullException.ThrowIfNull(context);
        if (formula.Trim().Length == 0)
        {
            throw new ArgumentException("The formula must not be empty.", nameof(formula));
        }

        if (formula.Length > _options.MaximumFormulaLength)
        {
            throw new ArgumentException(
                $"The formula is longer than the allowed {_options.MaximumFormulaLength} characters.",
                nameof(formula));
        }

        List<FormulaToken> tokens = FormulaTokenizer.Tokenize(formula);
        return new FormulaEvaluator(tokens, context).Evaluate();
    }
}
