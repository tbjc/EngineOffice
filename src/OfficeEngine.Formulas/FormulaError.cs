namespace OfficeEngine.Formulas;

/// <summary>The error values a formula can produce, the same ones Excel uses.</summary>
public sealed class FormulaError
{
    private FormulaError(string code, string name)
    {
        Code = code;
        Name = name;
    }

    /// <summary>Gets the code of the error, for example <c>#DIV/0!</c>.</summary>
    public string Code { get; }

    /// <summary>Gets the name of the error, for example <c>DivideByZero</c>.</summary>
    public string Name { get; }

    /// <summary>Gets the error of a division by zero.</summary>
    public static FormulaError DivideByZero { get; } = new("#DIV/0!", nameof(DivideByZero));

    /// <summary>Gets the error of a value of the wrong type.</summary>
    public static FormulaError Value { get; } = new("#VALUE!", nameof(Value));

    /// <summary>Gets the error of a reference that does not exist.</summary>
    public static FormulaError Reference { get; } = new("#REF!", nameof(Reference));

    /// <summary>Gets the error of an unknown function or name.</summary>
    public static FormulaError UnknownName { get; } = new("#NAME?", nameof(UnknownName));

    /// <summary>Gets the error of a missing value.</summary>
    public static FormulaError NotAvailable { get; } = new("#N/A", nameof(NotAvailable));

    /// <summary>Gets the error of a number that is out of range.</summary>
    public static FormulaError Number { get; } = new("#NUM!", nameof(Number));

    /// <summary>Gets the error of an intersection that does not exist.</summary>
    public static FormulaError Null { get; } = new("#NULL!", nameof(Null));

    /// <summary>Returns the error with the given code, or <see langword="null"/> when the code is unknown.</summary>
    /// <param name="code">The code of the error, for example <c>#DIV/0!</c>.</param>
    /// <returns>The error, or <see langword="null"/> when the code is not one of the known ones.</returns>
    public static FormulaError? FromCode(string code)
    {
        foreach (FormulaError error in All)
        {
            if (string.Equals(error.Code, code, StringComparison.OrdinalIgnoreCase))
            {
                return error;
            }
        }

        return null;
    }

    /// <summary>Gets every error value, in a stable order.</summary>
    public static IReadOnlyList<FormulaError> All { get; } =
    [
        DivideByZero,
        Value,
        Reference,
        UnknownName,
        NotAvailable,
        Number,
        Null,
    ];

    /// <inheritdoc />
    public override string ToString() => Code;
}
