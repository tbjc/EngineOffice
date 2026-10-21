using System.Globalization;

namespace OfficeEngine;

/// <summary>
/// An immutable 8-bit RGB colour used to describe explicit colours in Office documents.
/// </summary>
/// <remarks>
/// <para>
/// Office Open XML stores an explicit colour as a six-digit hexadecimal string such as
/// <c>FF0000</c>. Since such a value is always exactly three bytes wide it maps to this type
/// without loss, so reads and writes through the public API are reversible.
/// </para>
/// <para>
/// This type does not model theme colours or the <c>auto</c> keyword, which have no fixed RGB
/// value. Those cases are reported as "no explicit colour" by the Word API and their original
/// markup is preserved.
/// </para>
/// </remarks>
public readonly struct RgbColor : IEquatable<RgbColor>
{
    /// <summary>Initializes a new instance of the <see cref="RgbColor"/> struct.</summary>
    /// <param name="red">The red component.</param>
    /// <param name="green">The green component.</param>
    /// <param name="blue">The blue component.</param>
    public RgbColor(byte red, byte green, byte blue)
    {
        Red = red;
        Green = green;
        Blue = blue;
    }

    /// <summary>Gets the red component.</summary>
    public byte Red { get; }

    /// <summary>Gets the green component.</summary>
    public byte Green { get; }

    /// <summary>Gets the blue component.</summary>
    public byte Blue { get; }

    /// <summary>Creates a colour from its red, green and blue components.</summary>
    /// <param name="red">The red component.</param>
    /// <param name="green">The green component.</param>
    /// <param name="blue">The blue component.</param>
    /// <returns>The colour.</returns>
    public static RgbColor FromRgb(byte red, byte green, byte blue) => new(red, green, blue);

    /// <summary>
    /// Formats the colour as an upper-case six-digit hexadecimal string, for example <c>FF0000</c>.
    /// </summary>
    /// <returns>The hexadecimal representation, without a leading <c>#</c>.</returns>
    public string ToHex() => string.Concat(
        Red.ToString("X2", CultureInfo.InvariantCulture),
        Green.ToString("X2", CultureInfo.InvariantCulture),
        Blue.ToString("X2", CultureInfo.InvariantCulture));

    /// <summary>
    /// Tries to parse a six-digit hexadecimal string into a colour.
    /// </summary>
    /// <param name="hex">
    /// The string to parse. A leading <c>#</c> and surrounding whitespace are allowed.
    /// </param>
    /// <param name="color">
    /// When this method returns <see langword="true"/>, the parsed colour; otherwise, the default colour.
    /// </param>
    /// <returns><see langword="true"/> if the string could be parsed; otherwise, <see langword="false"/>.</returns>
    public static bool TryParseHex(string? hex, out RgbColor color)
    {
        color = default;
        if (hex is null)
        {
            return false;
        }

        ReadOnlySpan<char> span = hex.AsSpan().Trim();
        if (span.Length == 7 && span[0] == '#')
        {
            span = span[1..];
        }

        if (span.Length != 6
            || !byte.TryParse(span[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte red)
            || !byte.TryParse(span.Slice(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte green)
            || !byte.TryParse(span.Slice(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte blue))
        {
            return false;
        }

        color = new RgbColor(red, green, blue);
        return true;
    }

    /// <inheritdoc />
    public bool Equals(RgbColor other) => Red == other.Red && Green == other.Green && Blue == other.Blue;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is RgbColor other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(Red, Green, Blue);

    /// <summary>Determines whether two colours are equal.</summary>
    /// <param name="left">The first colour.</param>
    /// <param name="right">The second colour.</param>
    /// <returns><see langword="true"/> if both colours are equal; otherwise, <see langword="false"/>.</returns>
    public static bool operator ==(RgbColor left, RgbColor right) => left.Equals(right);

    /// <summary>Determines whether two colours are different.</summary>
    /// <param name="left">The first colour.</param>
    /// <param name="right">The second colour.</param>
    /// <returns><see langword="true"/> if both colours are different; otherwise, <see langword="false"/>.</returns>
    public static bool operator !=(RgbColor left, RgbColor right) => !left.Equals(right);

    /// <inheritdoc />
    public override string ToString() => $"#{ToHex()}";
}
