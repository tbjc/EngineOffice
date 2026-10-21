using System.Diagnostics.CodeAnalysis;

namespace OfficeEngine.Word;

/// <summary>
/// Alignment of a tab stop, mirroring a subset of OOXML <c>ST_TabStop</c>.
/// </summary>
/// <remarks>
/// Office Open XML also defines the values <c>start</c>, <c>end</c> and <c>num</c>, which this
/// version does not model. A paragraph that stores one of them is reported through
/// <see cref="UnsupportedOfficeFeatureException"/> instead of being guessed or silently ignored.
/// </remarks>
public enum TabStopAlignment
{
    /// <summary>Clears a tab stop inherited from a style or the document defaults.</summary>
    Clear,

    /// <summary>Left-aligned tab stop: text starts at the position.</summary>
    Left,

    /// <summary>Centered tab stop: text is centred on the position.</summary>
    Center,

    /// <summary>Right-aligned tab stop: text ends at the position.</summary>
    Right,

    /// <summary>Decimal tab stop (text is aligned on its decimal separator).</summary>
    [SuppressMessage(
        "Naming",
        "CA1720:Identifier contains type name",
        Justification = "The member mirrors the Office Open XML value 'decimal'; renaming it would hide that mapping.")]
    Decimal,

    /// <summary>Bar tab stop, which draws a vertical line at the position instead of aligning text.</summary>
    Bar,
}
