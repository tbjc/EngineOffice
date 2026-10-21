using System.Diagnostics.CodeAnalysis;

namespace OfficeEngine.Word;

/// <summary>
/// The numbering format of a list level (<c>w:numFmt</c>), mirroring the subset of
/// <c>ST_NumberFormat</c> that this version can create and read.
/// </summary>
/// <remarks>
/// <para>
/// Office Open XML defines about sixty number formats. This enum models the ones a caller needs in
/// order to build an ordinary bulleted or numbered list; the rest are still read as markup and, when
/// a document uses one, reading a level's format throws <see cref="UnsupportedOfficeFeatureException"/>
/// rather than reporting a wrong value.
/// </para>
/// <para>
/// The member names follow the Office Open XML spelling so they can be matched to <c>w:numFmt</c>
/// values at a glance.
/// </para>
/// </remarks>
public enum WordListFormat
{
    /// <summary>A bulleted list level (<c>bullet</c>).</summary>
    Bullet,

    /// <summary>Arabic numerals: 1, 2, 3 (<c>decimal</c>).</summary>
    [SuppressMessage(
        "Naming",
        "CA1720:Identifier contains type name",
        Justification = "The member mirrors the Office Open XML value 'decimal'; renaming it would hide that mapping.")]
    Decimal,

    /// <summary>Lowercase Latin letters: a, b, c (<c>lowerLetter</c>).</summary>
    LowerLetter,

    /// <summary>Uppercase Latin letters: A, B, C (<c>upperLetter</c>).</summary>
    UpperLetter,

    /// <summary>Lowercase Roman numerals: i, ii, iii (<c>lowerRoman</c>).</summary>
    LowerRoman,

    /// <summary>Uppercase Roman numerals: I, II, III (<c>upperRoman</c>).</summary>
    UpperRoman,
}
