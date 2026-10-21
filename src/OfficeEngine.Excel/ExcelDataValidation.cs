using System.Globalization;

namespace OfficeEngine.Excel;

/// <summary>
/// A data validation rule: the range of cells it applies to, the rule itself and the messages shown
/// when the rule is not met.
/// </summary>
/// <remarks>
/// <para>
/// This version creates list, whole number and decimal rules. A rule read from another workbook can
/// be of any kind; the ones that are not modelled are reported with their type but are not
/// interpreted.
/// </para>
/// <para>
/// The range of a rule read from a workbook that applies it to several areas reports the first area;
/// the rest of the areas are preserved in the file.
/// </para>
/// </remarks>
public sealed class ExcelDataValidation
{
    internal ExcelDataValidation()
    {
    }

    /// <summary>Gets the A1-style range the rule applies to, for example <c>A1:A10</c>.</summary>
    public string Range { get; internal init; } = string.Empty;

    /// <summary>Gets the kind of rule.</summary>
    public ExcelDataValidationType Type { get; internal init; } = ExcelDataValidationType.None;

    /// <summary>Gets the first formula of the rule, exactly as it is stored.</summary>
    public string? Formula1 { get; internal init; }

    /// <summary>Gets the second formula of the rule, exactly as it is stored.</summary>
    public string? Formula2 { get; internal init; }

    /// <summary>Gets the smallest value that a numeric rule accepts.</summary>
    public double? Minimum { get; internal init; }

    /// <summary>Gets the largest value that a numeric rule accepts.</summary>
    public double? Maximum { get; internal init; }

    /// <summary>Gets the values that a list rule accepts. It is empty for the rules that are not lists.</summary>
    public IReadOnlyList<string> AllowedValues { get; internal init; } = [];

    /// <summary>Gets whether a blank cell is accepted.</summary>
    public bool AllowBlank { get; internal init; }

    /// <summary>Gets the title of the error message.</summary>
    public string? ErrorTitle { get; internal init; }

    /// <summary>Gets the error message.</summary>
    public string? Error { get; internal init; }

    /// <summary>Gets the title of the input message.</summary>
    public string? PromptTitle { get; internal init; }

    /// <summary>Gets the input message.</summary>
    public string? Prompt { get; internal init; }

    /// <summary>Creates a rule that accepts only the given values.</summary>
    /// <param name="rangeReference">
    /// The A1-style range the rule applies to, for example <c>A1:A10</c>. A single cell is also
    /// accepted.
    /// </param>
    /// <param name="values">The accepted values. A value cannot contain a comma or a double quote.</param>
    /// <param name="allowBlank">Whether a blank cell is accepted.</param>
    /// <returns>The rule.</returns>
    /// <exception cref="ArgumentException">
    /// The range is not a valid A1-style reference, the list is empty or a value cannot be stored in
    /// an inline list.
    /// </exception>
    public static ExcelDataValidation List(string rangeReference, IEnumerable<string> values, bool allowBlank = true)
    {
        ArgumentException.ThrowIfNullOrEmpty(rangeReference);
        ArgumentNullException.ThrowIfNull(values);

        List<string> allowed = [];
        foreach (string value in values)
        {
            if (string.IsNullOrEmpty(value))
            {
                throw new ArgumentException("A value of the list must not be empty.", nameof(values));
            }

            if (value.Contains(',') || value.Contains('"'))
            {
                throw new ArgumentException(
                    "A value of the list cannot contain a comma or a double quote.",
                    nameof(values));
            }

            allowed.Add(value);
        }

        if (allowed.Count == 0)
        {
            throw new ArgumentException("The list must contain at least one value.", nameof(values));
        }

        return new ExcelDataValidation
        {
            Range = ExcelDataValidations.NormalizeRange(rangeReference),
            Type = ExcelDataValidationType.List,
            Formula1 = "\"" + string.Join(",", allowed) + "\"",
            AllowedValues = allowed,
            AllowBlank = allowBlank,
        };
    }

    /// <summary>Creates a rule that accepts only whole numbers.</summary>
    /// <param name="rangeReference">The A1-style range the rule applies to.</param>
    /// <param name="minimum">
    /// The smallest accepted value. When it is given and <paramref name="maximum"/> is not, the rule
    /// accepts the values that are greater than or equal to it.
    /// </param>
    /// <param name="maximum">
    /// The largest accepted value. When it is given and <paramref name="minimum"/> is not, the rule
    /// accepts the values that are less than or equal to it.
    /// </param>
    /// <param name="allowBlank">Whether a blank cell is accepted.</param>
    /// <returns>The rule.</returns>
    /// <exception cref="ArgumentException">
    /// The range is not a valid A1-style reference or the minimum is greater than the maximum.
    /// </exception>
    public static ExcelDataValidation WholeNumber(
        string rangeReference,
        double? minimum = null,
        double? maximum = null,
        bool allowBlank = true)
        => Create(ExcelDataValidationType.WholeNumber, rangeReference, minimum, maximum, allowBlank);

    /// <summary>Creates a rule that accepts only numbers with decimals.</summary>
    /// <param name="rangeReference">The A1-style range the rule applies to.</param>
    /// <param name="minimum">The smallest accepted value.</param>
    /// <param name="maximum">The largest accepted value.</param>
    /// <param name="allowBlank">Whether a blank cell is accepted.</param>
    /// <returns>The rule.</returns>
    /// <exception cref="ArgumentException">
    /// The range is not a valid A1-style reference or the minimum is greater than the maximum.
    /// </exception>
    public static ExcelDataValidation DecimalNumber(
        string rangeReference,
        double? minimum = null,
        double? maximum = null,
        bool allowBlank = true)
        => Create(ExcelDataValidationType.DecimalNumber, rangeReference, minimum, maximum, allowBlank);

    /// <inheritdoc />
    public override string ToString() => $"{Range}: {Type}";

    private static ExcelDataValidation Create(
        ExcelDataValidationType type,
        string rangeReference,
        double? minimum,
        double? maximum,
        bool allowBlank)
    {
        if (minimum is double lower && maximum is double upper && lower > upper)
        {
            throw new ArgumentException("The minimum cannot be greater than the maximum.", nameof(minimum));
        }

        string? formula1 = minimum?.ToString("R", CultureInfo.InvariantCulture);
        string? formula2 = maximum?.ToString("R", CultureInfo.InvariantCulture);
        return new ExcelDataValidation
        {
            Range = ExcelDataValidations.NormalizeRange(rangeReference),
            Type = type,
            Formula1 = formula1,
            Formula2 = formula2,
            Minimum = minimum,
            Maximum = maximum,
            AllowBlank = allowBlank,
        };
    }
}
