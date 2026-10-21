namespace OfficeEngine.Excel;

/// <summary>One edge of a cell border: its line style and optional RGB colour.</summary>
/// <remarks>
/// A <see cref="Color"/> of <see langword="null"/> means "no explicit colour", which is how this
/// version reports the <c>auto</c> keyword and theme colours.
/// </remarks>
public readonly struct CellBorder : IEquatable<CellBorder>
{
    /// <summary>Initializes a new instance of the <see cref="CellBorder"/> struct.</summary>
    /// <param name="style">The line style of the edge.</param>
    /// <param name="color">The colour of the line, or <see langword="null"/> for no explicit colour.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="style"/> is not a defined value.</exception>
    public CellBorder(CellBorderStyle style, RgbColor? color = null)
    {
        if (!Enum.IsDefined(style))
        {
            throw new ArgumentOutOfRangeException(nameof(style), style, "The cell border style is not a defined value.");
        }

        Style = style;
        Color = color;
    }

    /// <summary>Gets the line style of the edge.</summary>
    public CellBorderStyle Style { get; }

    /// <summary>
    /// Gets the colour of the edge, or <see langword="null"/> when it has no explicit RGB colour.
    /// </summary>
    public RgbColor? Color { get; }

    /// <inheritdoc />
    public bool Equals(CellBorder other) => Style == other.Style && Color == other.Color;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is CellBorder other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(Style, Color);

    /// <summary>Determines whether two borders are equal.</summary>
    /// <param name="left">The first border.</param>
    /// <param name="right">The second border.</param>
    /// <returns><see langword="true"/> if both borders are equal; otherwise, <see langword="false"/>.</returns>
    public static bool operator ==(CellBorder left, CellBorder right) => left.Equals(right);

    /// <summary>Determines whether two borders are different.</summary>
    /// <param name="left">The first border.</param>
    /// <param name="right">The second border.</param>
    /// <returns><see langword="true"/> if both borders are different; otherwise, <see langword="false"/>.</returns>
    public static bool operator !=(CellBorder left, CellBorder right) => !left.Equals(right);
}
