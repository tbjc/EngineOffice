using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// One border edge of a table or of a cell: the line style, its width in points and its colour.
/// </summary>
/// <remarks>
/// <para>
/// Office Open XML stores every edge (<c>w:top</c>, <c>w:left</c>, <c>w:insideH</c>, …) as a
/// <c>w:val</c> with the line style, a <c>w:sz</c> with the width in eighths of a point and a
/// <c>w:color</c>. This type keeps those three values. The width is reported in points, so it is
/// read and written with a resolution of <c>0.125</c> points; an edge that stores no <c>w:sz</c>
/// reports the width Office Open XML uses by default, half a point. Attributes this version does
/// not model — <c>w:space</c>, a theme colour — are preserved when the edge is rewritten.
/// </para>
/// <para>
/// A <see cref="Color"/> of <see langword="null"/> means "no explicit colour", which is how Office
/// Open XML expresses both the <c>auto</c> keyword and theme colours, as elsewhere in this API.
/// </para>
/// </remarks>
public readonly struct TableBorder : IEquatable<TableBorder>
{
    private const double DefaultSizePoints = 0.5;

    // The non-decorative values of ST_Border, keyed by the string Office Open XML stores. The art
    // borders (apples, balloons, hearts…) are deliberately absent: reading one is unsupported.
    private static readonly Dictionary<string, BorderStyle> StylesFromOoxml = new(StringComparer.Ordinal)
    {
        ["nil"] = BorderStyle.Nil,
        ["none"] = BorderStyle.None,
        ["single"] = BorderStyle.SingleLine,
        ["thick"] = BorderStyle.Thick,
        ["double"] = BorderStyle.DoubleLine,
        ["dotted"] = BorderStyle.Dotted,
        ["dashed"] = BorderStyle.Dashed,
        ["dotDash"] = BorderStyle.DotDash,
        ["dotDotDash"] = BorderStyle.DotDotDash,
        ["triple"] = BorderStyle.Triple,
        ["thinThickSmallGap"] = BorderStyle.ThinThickSmallGap,
        ["thickThinSmallGap"] = BorderStyle.ThickThinSmallGap,
        ["thinThickThinSmallGap"] = BorderStyle.ThinThickThinSmallGap,
        ["thinThickMediumGap"] = BorderStyle.ThinThickMediumGap,
        ["thickThinMediumGap"] = BorderStyle.ThickThinMediumGap,
        ["thinThickThinMediumGap"] = BorderStyle.ThinThickThinMediumGap,
        ["thinThickLargeGap"] = BorderStyle.ThinThickLargeGap,
        ["thickThinLargeGap"] = BorderStyle.ThickThinLargeGap,
        ["thinThickThinLargeGap"] = BorderStyle.ThinThickThinLargeGap,
        ["wave"] = BorderStyle.Wave,
        ["doubleWave"] = BorderStyle.DoubleWave,
        ["dashSmallGap"] = BorderStyle.DashSmallGap,
        ["dashDotStroked"] = BorderStyle.DashDotStroked,
        ["threeDEmboss"] = BorderStyle.ThreeDEmboss,
        ["threeDEngrave"] = BorderStyle.ThreeDEngrave,
        ["outset"] = BorderStyle.Outset,
        ["inset"] = BorderStyle.Inset,
    };

    // The same table read backwards, so writing a style cannot drift from reading it.
    private static readonly Dictionary<BorderStyle, string> StylesToOoxml =
        StylesFromOoxml.ToDictionary(pair => pair.Value, pair => pair.Key);

    /// <summary>Initializes a new instance of the <see cref="TableBorder"/> struct.</summary>
    /// <param name="style">The line style of the edge.</param>
    /// <param name="sizePoints">The width of the line, in points. Must be finite and non-negative.</param>
    /// <param name="color">
    /// The colour of the line, or <see langword="null"/> to leave it explicit-less (OOXML <c>auto</c>).
    /// </param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="style"/> is not a defined style, or <paramref name="sizePoints"/> is not a
    /// finite, non-negative number, or it is too large to be represented in eighths of a point.
    /// </exception>
    public TableBorder(BorderStyle style, double sizePoints, RgbColor? color = null)
    {
        if (!Enum.IsDefined(style))
        {
            throw new ArgumentOutOfRangeException(nameof(style), style, "The border style is not a defined value.");
        }

        if (!double.IsFinite(sizePoints) || sizePoints < 0.0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sizePoints),
                sizePoints,
                "The border width must be a finite, non-negative number of points.");
        }

        double eighths = Math.Round(sizePoints * 8.0, MidpointRounding.AwayFromZero);
        if (eighths > uint.MaxValue)
        {
            throw new ArgumentOutOfRangeException(
                nameof(sizePoints),
                sizePoints,
                "The border width is too large to be represented in eighths of a point.");
        }

        Style = style;
        SizePoints = eighths / 8.0;
        Color = color;
    }

    /// <summary>Gets the line style of the edge.</summary>
    public BorderStyle Style { get; }

    /// <summary>Gets the width of the line, in points.</summary>
    public double SizePoints { get; }

    /// <summary>
    /// Gets the colour of the line, or <see langword="null"/> when the edge has no explicit colour.
    /// </summary>
    public RgbColor? Color { get; }

    /// <summary>
    /// Reads the border an edge element stores.
    /// </summary>
    /// <param name="edge">The <c>w:top</c>, <c>w:tcBorders/w:left</c> or equivalent element.</param>
    /// <returns>The stored border.</returns>
    /// <exception cref="InvalidOfficeFileException">A stored value cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The edge uses a line style that Office Open XML defines but this version does not model.
    /// </exception>
    internal static TableBorder Read(W.BorderType edge)
    {
        string? style = edge.Val?.InnerText;
        if (style is null)
        {
            throw new InvalidOfficeFileException("A table border stores no w:val attribute, which this document cannot read.");
        }

        if (!StylesFromOoxml.TryGetValue(style, out BorderStyle mapped))
        {
            throw new UnsupportedOfficeFeatureException(
                $"A table border uses the style '{style}', which this version does not model.");
        }

        return new TableBorder(mapped, ReadSize(edge.Size), ReadColor(edge.Color));
    }

    /// <summary>
    /// Writes this border into the edge element that stores it, leaving any attribute this version
    /// does not model untouched.
    /// </summary>
    /// <param name="edge">The element to write.</param>
    internal void ApplyTo(W.BorderType edge)
    {
        // Map the style before the first write so that an undefined value leaves the markup untouched.
        W.BorderValues style = ToOoxml(Style);

        edge.Val = style;
        edge.Size = (uint)Math.Round(SizePoints * 8.0, MidpointRounding.AwayFromZero);
        edge.Color = Color?.ToHex() ?? "auto";
    }

    /// <inheritdoc />
    public bool Equals(TableBorder other)
        => Style == other.Style && SizePoints.Equals(other.SizePoints) && Nullable.Equals(Color, other.Color);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is TableBorder other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(Style, SizePoints, Color);

    /// <summary>Determines whether two borders are equal.</summary>
    /// <param name="left">The first border.</param>
    /// <param name="right">The second border.</param>
    /// <returns><see langword="true"/> if both borders are equal; otherwise, <see langword="false"/>.</returns>
    public static bool operator ==(TableBorder left, TableBorder right) => left.Equals(right);

    /// <summary>Determines whether two borders are different.</summary>
    /// <param name="left">The first border.</param>
    /// <param name="right">The second border.</param>
    /// <returns><see langword="true"/> if both borders are different; otherwise, <see langword="false"/>.</returns>
    public static bool operator !=(TableBorder left, TableBorder right) => !left.Equals(right);

    /// <inheritdoc />
    public override string ToString()
        => string.Create(CultureInfo.InvariantCulture, $"{Style} {SizePoints} pt{(Color is { } color ? $" {color}" : string.Empty)}");

    private static W.BorderValues ToOoxml(BorderStyle style) => new(StylesToOoxml[style]);

    private static double ReadSize(UInt32Value? stored)
    {
        if (stored is null)
        {
            return DefaultSizePoints;
        }

        if (!uint.TryParse(stored.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint eighths))
        {
            throw new InvalidOfficeFileException(
                $"A table border stores the width '{stored.InnerText}' that this document cannot read.");
        }

        return eighths / 8.0;
    }

    private static RgbColor? ReadColor(StringValue? stored)
    {
        if (stored is null)
        {
            return null;
        }

        string? value = stored.Value;
        if (value is null || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (!RgbColor.TryParseHex(value, out RgbColor color))
        {
            throw new InvalidOfficeFileException($"A table border stores the colour '{value}' that this document cannot read.");
        }

        return color;
    }
}
