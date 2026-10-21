using System;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// The page size and orientation of a section, as <c>w:sectPr/w:pgSz</c> stores them.
/// </summary>
/// <remarks>
/// <para>
/// Office Open XML stores the page size in twentieths of a point (twips), so a value is read and
/// written with a resolution of <c>0.05</c> points. The orientation is stored separately in
/// <c>w:orient</c> and does not change the dimensions by itself: use <see cref="Rotate"/> to swap
/// them when switching between portrait and landscape.
/// </para>
/// <para>
/// Both dimensions are required by the schema. A document that stores a <c>w:pgSz</c> without one of
/// them fails with <see cref="InvalidOfficeFileException"/> instead of inventing a size.
/// </para>
/// </remarks>
public readonly struct PageSize : IEquatable<PageSize>
{
    private readonly uint _widthTwips;
    private readonly uint _heightTwips;

    /// <summary>Creates a page size from its dimensions and orientation.</summary>
    /// <param name="widthPoints">The page width, in points. Must be finite and positive.</param>
    /// <param name="heightPoints">The page height, in points. Must be finite and positive.</param>
    /// <param name="orientation">The page orientation.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A dimension is not a finite, positive number, or it is too large to be represented in twips
    /// (twentieths of a point); or <paramref name="orientation"/> is not a defined value.
    /// </exception>
    public PageSize(double widthPoints, double heightPoints, PageOrientation orientation)
        : this(ToTwips(widthPoints, nameof(widthPoints)), ToTwips(heightPoints, nameof(heightPoints)), orientation)
    {
    }

    private PageSize(uint widthTwips, uint heightTwips, PageOrientation orientation)
    {
        if (orientation is not (PageOrientation.Portrait or PageOrientation.Landscape))
        {
            throw new ArgumentOutOfRangeException(
                nameof(orientation),
                orientation,
                "The page orientation is not a defined value.");
        }

        _widthTwips = widthTwips;
        _heightTwips = heightTwips;
        Orientation = orientation;
    }

    /// <summary>Gets the US Letter page size (8.5 × 11 inches), in portrait orientation.</summary>
    public static PageSize Letter { get; } = new(12240, 15840, PageOrientation.Portrait);

    /// <summary>Gets the ISO A4 page size (210 × 297 mm), in portrait orientation.</summary>
    public static PageSize A4 { get; } = new(11906, 16838, PageOrientation.Portrait);

    /// <summary>Gets the page width in points.</summary>
    public double WidthPoints => _widthTwips / 20.0;

    /// <summary>Gets the page height in points.</summary>
    public double HeightPoints => _heightTwips / 20.0;

    /// <summary>Gets the orientation stored in <c>w:orient</c>.</summary>
    public PageOrientation Orientation { get; }

    /// <summary>
    /// Returns the same page size with its dimensions swapped and its orientation toggled.
    /// </summary>
    /// <returns>The rotated page size.</returns>
    public PageSize Rotate()
        => new(
            _heightTwips,
            _widthTwips,
            Orientation == PageOrientation.Portrait ? PageOrientation.Landscape : PageOrientation.Portrait);

    internal uint WidthTwips => _widthTwips;

    internal uint HeightTwips => _heightTwips;

    internal static PageSize? Read(W.PageSize? stored, string propertyDescription)
    {
        if (stored is null)
        {
            return null;
        }

        uint width = ReadMeasure(stored.Width, propertyDescription);
        uint height = ReadMeasure(stored.Height, propertyDescription);

        PageOrientation orientation = PageOrientation.Portrait;
        if (stored.Orient is not null)
        {
            W.PageOrientationValues value = ReadStoredValue(() => stored.Orient.Value, propertyDescription);
            orientation = value == W.PageOrientationValues.Landscape
                ? PageOrientation.Landscape
                : value == W.PageOrientationValues.Portrait
                    ? PageOrientation.Portrait
                    : throw new UnsupportedOfficeFeatureException(
                        $"A {propertyDescription} uses the orientation '{value}', a value that Office Open XML defines but this version does not model.");
        }

        return new PageSize(width, height, orientation);
    }

    internal void ApplyTo(W.PageSize element)
    {
        element.Width = _widthTwips;
        element.Height = _heightTwips;
        element.Orient = Orientation == PageOrientation.Landscape
            ? W.PageOrientationValues.Landscape
            : W.PageOrientationValues.Portrait;
    }

    private static uint ToTwips(double points, string parameterName)
    {
        if (!double.IsFinite(points) || points <= 0.0)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The page dimension must be a finite, positive number of points.");
        }

        double twips = Math.Round(points * 20.0, MidpointRounding.AwayFromZero);
        if (twips < 1.0 || twips > uint.MaxValue)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The page dimension must be at least 0.05 points and fit in twips (twentieths of a point).");
        }

        return (uint)twips;
    }

    private static uint ReadMeasure(DocumentFormat.OpenXml.UInt32Value? stored, string propertyDescription)
    {
        if (stored is null)
        {
            throw new InvalidOfficeFileException(
                $"A {propertyDescription} stores a w:pgSz element without both required dimensions.");
        }

        try
        {
            return stored.Value;
        }
        catch (FormatException exception)
        {
            throw new InvalidOfficeFileException(
                $"A {propertyDescription} stores the page dimension '{stored.InnerText}' that this document cannot read.",
                exception);
        }
    }

    private static T ReadStoredValue<T>(Func<T> read, string propertyDescription)
    {
        try
        {
            return read();
        }
        catch (FormatException exception)
        {
            throw new InvalidOfficeFileException(
                $"A {propertyDescription} stores a page orientation that this document cannot read.",
                exception);
        }
    }

    /// <inheritdoc />
    public bool Equals(PageSize other)
        => _widthTwips == other._widthTwips && _heightTwips == other._heightTwips && Orientation == other.Orientation;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is PageSize other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(_widthTwips, _heightTwips, Orientation);

    /// <summary>Determines whether two page sizes are equal.</summary>
    /// <param name="left">The first page size.</param>
    /// <param name="right">The second page size.</param>
    /// <returns><see langword="true"/> if both are equal; otherwise, <see langword="false"/>.</returns>
    public static bool operator ==(PageSize left, PageSize right) => left.Equals(right);

    /// <summary>Determines whether two page sizes are different.</summary>
    /// <param name="left">The first page size.</param>
    /// <param name="right">The second page size.</param>
    /// <returns><see langword="true"/> if both are different; otherwise, <see langword="false"/>.</returns>
    public static bool operator !=(PageSize left, PageSize right) => !left.Equals(right);

    /// <inheritdoc />
    public override string ToString()
        => string.Create(
            CultureInfo.InvariantCulture,
            $"{WidthPoints} × {HeightPoints} pt ({Orientation})");
}
