using System;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// The page margins of a section, as <c>w:sectPr/w:pgMar</c> stores them.
/// </summary>
/// <remarks>
/// <para>
/// Office Open XML stores the margins in twentieths of a point (twips), so a value is read and
/// written with a resolution of <c>0.05</c> points. The top and bottom margins are signed by the
/// schema; the right and left margins must not be negative.
/// </para>
/// <para>
/// The schema requires all four margins when <c>w:pgMar</c> exists. A document that omits one of them
/// fails with <see cref="InvalidOfficeFileException"/> instead of inventing a margin. The header,
/// footer and gutter distances are not modelled here; they are preserved when the margins are
/// edited.
/// </para>
/// </remarks>
public readonly struct PageMargins : IEquatable<PageMargins>
{
    private readonly int _topTwips;
    private readonly uint _rightTwips;
    private readonly int _bottomTwips;
    private readonly uint _leftTwips;

    /// <summary>Creates page margins from their measures in points.</summary>
    /// <param name="topPoints">The top margin, in points.</param>
    /// <param name="rightPoints">The right margin, in points. Must not be negative.</param>
    /// <param name="bottomPoints">The bottom margin, in points.</param>
    /// <param name="leftPoints">The left margin, in points. Must not be negative.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A value is not finite, a horizontal margin is negative, or a value is too large to be
    /// represented in twips (twentieths of a point).
    /// </exception>
    public PageMargins(double topPoints, double rightPoints, double bottomPoints, double leftPoints)
        : this(
            ToSignedTwips(topPoints, nameof(topPoints)),
            ToUnsignedTwips(rightPoints, nameof(rightPoints)),
            ToSignedTwips(bottomPoints, nameof(bottomPoints)),
            ToUnsignedTwips(leftPoints, nameof(leftPoints)))
    {
    }

    private PageMargins(int topTwips, uint rightTwips, int bottomTwips, uint leftTwips)
    {
        _topTwips = topTwips;
        _rightTwips = rightTwips;
        _bottomTwips = bottomTwips;
        _leftTwips = leftTwips;
    }

    /// <summary>Gets the top margin in points.</summary>
    public double TopPoints => _topTwips / 20.0;

    /// <summary>Gets the right margin in points.</summary>
    public double RightPoints => _rightTwips / 20.0;

    /// <summary>Gets the bottom margin in points.</summary>
    public double BottomPoints => _bottomTwips / 20.0;

    /// <summary>Gets the left margin in points.</summary>
    public double LeftPoints => _leftTwips / 20.0;

    internal static PageMargins? Read(W.PageMargin? stored, string propertyDescription)
    {
        if (stored is null)
        {
            return null;
        }

        if (stored.Top is null || stored.Right is null || stored.Bottom is null || stored.Left is null)
        {
            throw new InvalidOfficeFileException(
                $"A {propertyDescription} stores a w:pgMar element without all four required margins.");
        }

        try
        {
            return new PageMargins(stored.Top.Value, stored.Right.Value, stored.Bottom.Value, stored.Left.Value);
        }
        catch (FormatException exception)
        {
            throw new InvalidOfficeFileException(
                $"A {propertyDescription} stores a page margin that this document cannot read.",
                exception);
        }
    }

    internal void ApplyTo(W.PageMargin element)
    {
        element.Top = _topTwips;
        element.Right = _rightTwips;
        element.Bottom = _bottomTwips;
        element.Left = _leftTwips;
    }

    private static int ToSignedTwips(double points, string parameterName)
    {
        if (!double.IsFinite(points))
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The page margin must be a finite number of points.");
        }

        double twips = Math.Round(points * 20.0, MidpointRounding.AwayFromZero);
        if (twips < int.MinValue || twips > int.MaxValue)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The page margin is too large to be represented in twips (twentieths of a point).");
        }

        return (int)twips;
    }

    private static uint ToUnsignedTwips(double points, string parameterName)
    {
        if (!double.IsFinite(points) || points < 0.0)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The horizontal page margin must be a finite, non-negative number of points.");
        }

        double twips = Math.Round(points * 20.0, MidpointRounding.AwayFromZero);
        if (twips > uint.MaxValue)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The page margin is too large to be represented in twips (twentieths of a point).");
        }

        return (uint)twips;
    }

    /// <inheritdoc />
    public bool Equals(PageMargins other)
        => _topTwips == other._topTwips
            && _rightTwips == other._rightTwips
            && _bottomTwips == other._bottomTwips
            && _leftTwips == other._leftTwips;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is PageMargins other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(_topTwips, _rightTwips, _bottomTwips, _leftTwips);

    /// <summary>Determines whether two sets of page margins are equal.</summary>
    /// <param name="left">The first margins.</param>
    /// <param name="right">The second margins.</param>
    /// <returns><see langword="true"/> if both are equal; otherwise, <see langword="false"/>.</returns>
    public static bool operator ==(PageMargins left, PageMargins right) => left.Equals(right);

    /// <summary>Determines whether two sets of page margins are different.</summary>
    /// <param name="left">The first margins.</param>
    /// <param name="right">The second margins.</param>
    /// <returns><see langword="true"/> if both are different; otherwise, <see langword="false"/>.</returns>
    public static bool operator !=(PageMargins left, PageMargins right) => !left.Equals(right);

    /// <inheritdoc />
    public override string ToString()
        => string.Create(
            CultureInfo.InvariantCulture,
            $"top {TopPoints} pt, right {RightPoints} pt, bottom {BottomPoints} pt, left {LeftPoints} pt");
}
