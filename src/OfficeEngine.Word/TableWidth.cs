using System;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// The preferred width of a table or of a cell, as <c>w:tblW</c> and <c>w:tcW</c> store it.
/// </summary>
/// <remarks>
/// <para>
/// Office Open XML stores a width as an integer plus a unit (<c>w:type</c>): <c>auto</c> (no fixed
/// width), <c>dxa</c> (twentieths of a point) or <c>pct</c> (fiftieths of a percent). This type
/// keeps the unit and reports the measure the way the rest of the API does it: points for
/// <see cref="TableWidthType.Points"/> and a percentage for <see cref="TableWidthType.Percentage"/>.
/// A width is therefore read and written with the resolution of its unit, <c>0.05</c> points or
/// <c>0.02</c> percent.
/// </para>
/// <para>
/// The unit <c>nil</c>, which Office Open XML also defines, is not modelled: reading a width that
/// uses it fails with <see cref="UnsupportedOfficeFeatureException"/> and the original markup is
/// preserved.
/// </para>
/// <para>
/// Fixed column widths are not modelled, and neither is the width of the table grid
/// (<c>w:tblGrid</c>): a width set here is the table's or the cell's own preferred width, which is
/// what Word reconciles with the grid when it lays the table out.
/// </para>
/// </remarks>
public readonly struct TableWidth : IEquatable<TableWidth>
{
    private readonly int _measure;

    private TableWidth(TableWidthType type, int measure)
    {
        Type = type;
        _measure = measure;
    }

    /// <summary>
    /// Gets a width that lets the layout algorithm decide (OOXML <c>auto</c>).
    /// </summary>
    public static TableWidth Auto { get; } = new(TableWidthType.Auto, 0);

    /// <summary>Gets the kind of measure this width uses.</summary>
    public TableWidthType Type { get; }

    /// <summary>
    /// Gets the measure: points for <see cref="TableWidthType.Points"/>, a percentage for
    /// <see cref="TableWidthType.Percentage"/>, and <c>0</c> for <see cref="TableWidthType.Auto"/>.
    /// </summary>
    public double Value => Type switch
    {
        TableWidthType.Points => _measure / 20.0,
        TableWidthType.Percentage => _measure / 50.0,
        _ => 0.0,
    };

    /// <summary>Creates a width expressed in points (OOXML <c>w:type="dxa"</c>).</summary>
    /// <param name="points">The width, in points. Must be finite and non-negative.</param>
    /// <returns>The width.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="points"/> is not a finite, non-negative number, or it is too large to be
    /// represented in twips (twentieths of a point).
    /// </exception>
    public static TableWidth Points(double points)
    {
        if (!double.IsFinite(points) || points < 0.0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(points),
                points,
                "The table width must be a finite, non-negative number of points.");
        }

        double twips = Math.Round(points * 20.0, MidpointRounding.AwayFromZero);
        if (twips > int.MaxValue)
        {
            throw new ArgumentOutOfRangeException(
                nameof(points),
                points,
                "The table width is too large to be represented in twips (twentieths of a point).");
        }

        return new TableWidth(TableWidthType.Points, (int)twips);
    }

    /// <summary>Creates a width expressed as a percentage (OOXML <c>w:type="pct"</c>).</summary>
    /// <param name="percent">The width, as a percentage. Must be finite and non-negative.</param>
    /// <returns>The width.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="percent"/> is not a finite, non-negative number, or it is too large to be
    /// represented in fiftieths of a percent.
    /// </exception>
    public static TableWidth Percentage(double percent)
    {
        if (!double.IsFinite(percent) || percent < 0.0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(percent),
                percent,
                "The table width must be a finite, non-negative percentage.");
        }

        double fiftieths = Math.Round(percent * 50.0, MidpointRounding.AwayFromZero);
        if (fiftieths > int.MaxValue)
        {
            throw new ArgumentOutOfRangeException(
                nameof(percent),
                percent,
                "The table width is too large to be represented in fiftieths of a percent.");
        }

        return new TableWidth(TableWidthType.Percentage, (int)fiftieths);
    }

    /// <summary>
    /// Reads the width an element stores.
    /// </summary>
    /// <param name="stored">The <c>w:tblW</c> or <c>w:tcW</c> element, or <see langword="null"/>.</param>
    /// <param name="propertyDescription">How the caller refers to the width in error messages.</param>
    /// <returns>The stored width, or <see langword="null"/> when the element is absent.</returns>
    /// <exception cref="InvalidOfficeFileException">The stored width cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The stored width uses a unit this version does not model.
    /// </exception>
    internal static TableWidth? Read(W.TableWidthType? stored, string propertyDescription)
    {
        if (stored is null)
        {
            return null;
        }

        string? unit = stored.Type?.InnerText;
        if (unit is null)
        {
            // ST_TblWidth: "If this attribute is omitted, then its value shall be assumed to be dxa."
            unit = "dxa";
        }

        if (string.Equals(unit, "auto", StringComparison.Ordinal))
        {
            return Auto;
        }

        if (string.Equals(unit, "nil", StringComparison.Ordinal))
        {
            throw new UnsupportedOfficeFeatureException(
                $"A {propertyDescription} uses the unit 'nil', a value that Office Open XML defines but this version does not model.");
        }

        if (!string.Equals(unit, "dxa", StringComparison.Ordinal)
            && !string.Equals(unit, "pct", StringComparison.Ordinal))
        {
            throw new InvalidOfficeFileException(
                $"A {propertyDescription} uses the unit '{unit}' that this document cannot read.");
        }

        string? measured = stored.Width?.Value;
        if (measured is null)
        {
            // w:w is optional in the schema; an omitted measure reads as zero.
            measured = "0";
        }

        if (!int.TryParse(measured, NumberStyles.Integer, CultureInfo.InvariantCulture, out int measure))
        {
            throw new InvalidOfficeFileException(
                $"A {propertyDescription} stores the measure '{measured}' that this document cannot read.");
        }

        if (measure < 0)
        {
            throw new InvalidOfficeFileException(
                $"A {propertyDescription} stores the negative measure '{measured}' that this document cannot read.");
        }

        return new TableWidth(
            string.Equals(unit, "dxa", StringComparison.Ordinal) ? TableWidthType.Points : TableWidthType.Percentage,
            measure);
    }

    /// <summary>
    /// Writes this width into the element that stores it, leaving any attribute this version does
    /// not model untouched.
    /// </summary>
    /// <param name="element">The <c>w:tblW</c> or <c>w:tcW</c> element to write.</param>
    internal void ApplyTo(W.TableWidthType element)
    {
        element.Width = _measure.ToString(CultureInfo.InvariantCulture);
        element.Type = Type switch
        {
            TableWidthType.Points => W.TableWidthUnitValues.Dxa,
            TableWidthType.Percentage => W.TableWidthUnitValues.Pct,
            _ => W.TableWidthUnitValues.Auto,
        };
    }

    /// <inheritdoc />
    public bool Equals(TableWidth other) => Type == other.Type && _measure == other._measure;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is TableWidth other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(Type, _measure);

    /// <summary>Determines whether two widths are equal.</summary>
    /// <param name="left">The first width.</param>
    /// <param name="right">The second width.</param>
    /// <returns><see langword="true"/> if both widths are equal; otherwise, <see langword="false"/>.</returns>
    public static bool operator ==(TableWidth left, TableWidth right) => left.Equals(right);

    /// <summary>Determines whether two widths are different.</summary>
    /// <param name="left">The first width.</param>
    /// <param name="right">The second width.</param>
    /// <returns><see langword="true"/> if both widths are different; otherwise, <see langword="false"/>.</returns>
    public static bool operator !=(TableWidth left, TableWidth right) => !left.Equals(right);

    /// <inheritdoc />
    public override string ToString() => Type switch
    {
        TableWidthType.Auto => "auto",
        TableWidthType.Points => string.Create(CultureInfo.InvariantCulture, $"{Value} pt"),
        _ => string.Create(CultureInfo.InvariantCulture, $"{Value} %"),
    };
}
