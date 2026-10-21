namespace OfficeEngine.Excel;

/// <summary>The margins of the printed page, in inches.</summary>
public sealed class ExcelPageMargins
{
    internal ExcelPageMargins()
    {
    }

    /// <summary>Initializes a new instance of the <see cref="ExcelPageMargins"/> class.</summary>
    /// <param name="left">The left margin.</param>
    /// <param name="right">The right margin.</param>
    /// <param name="top">The top margin.</param>
    /// <param name="bottom">The bottom margin.</param>
    /// <param name="header">The distance from the top edge to the header.</param>
    /// <param name="footer">The distance from the bottom edge to the footer.</param>
    /// <exception cref="ArgumentOutOfRangeException">A margin is negative or greater than 100 inches.</exception>
    public ExcelPageMargins(
        double left,
        double right,
        double top,
        double bottom,
        double header = 0.3,
        double footer = 0.3)
    {
        Left = Validate(left, nameof(left));
        Right = Validate(right, nameof(right));
        Top = Validate(top, nameof(top));
        Bottom = Validate(bottom, nameof(bottom));
        Header = Validate(header, nameof(header));
        Footer = Validate(footer, nameof(footer));
    }

    /// <summary>Gets the left margin, in inches.</summary>
    public double Left { get; internal init; }

    /// <summary>Gets the right margin, in inches.</summary>
    public double Right { get; internal init; }

    /// <summary>Gets the top margin, in inches.</summary>
    public double Top { get; internal init; }

    /// <summary>Gets the bottom margin, in inches.</summary>
    public double Bottom { get; internal init; }

    /// <summary>Gets the distance from the top edge to the header, in inches.</summary>
    public double Header { get; internal init; }

    /// <summary>Gets the distance from the bottom edge to the footer, in inches.</summary>
    public double Footer { get; internal init; }

    /// <inheritdoc />
    public override string ToString() => $"{Left}/{Right}/{Top}/{Bottom}";

    private static double Validate(double value, string parameterName)
    {
        if (!double.IsFinite(value) || value < 0d || value > 100d)
        {
            throw new ArgumentOutOfRangeException(
                parameterName, value, "A margin must be between 0 and 100 inches.");
        }

        return value;
    }
}
