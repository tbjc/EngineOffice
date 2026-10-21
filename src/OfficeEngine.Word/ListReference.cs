using System;
using DocumentFormat.OpenXml;

namespace OfficeEngine.Word;

/// <summary>
/// Helpers shared by the members that read or write a numbering reference
/// (<see cref="Paragraph.ListId"/>, <see cref="Paragraph.ListLevel"/>) and by the numbering
/// catalogue (<see cref="WordList"/>, <see cref="WordListLevel"/>).
/// </summary>
/// <remarks>
/// <para>
/// A list identifier (<c>w:numPr/w:numId</c>, <c>w:num/@w:numId</c>) is a non-negative decimal
/// number and a list level (<c>w:numPr/w:ilvl</c>) is one of the nine levels Office Open XML
/// allows, from <c>0</c> to <see cref="MaxLevel"/>. Neither is checked for referential integrity:
/// Office Open XML lets a paragraph point to a list that does not exist, and Word renders the
/// paragraph as if it were not numbered.
/// </para>
/// <para>
/// <see cref="TryReadInt"/> is deliberately lenient. It is only used to resolve a reference while
/// looking something up, where a value that cannot be read can only fail to match. The public
/// properties read strictly and throw <see cref="InvalidOfficeFileException"/> instead.
/// </para>
/// </remarks>
internal static class ListReference
{
    /// <summary>The highest list level the schema allows (levels run from <c>0</c> to <c>8</c>).</summary>
    internal const int MaxLevel = 8;

    /// <summary>
    /// Throws when a list identifier is negative.
    /// </summary>
    /// <param name="value">The identifier to check.</param>
    /// <param name="parameterName">The parameter the identifier came from.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="value"/> is negative.</exception>
    internal static void EnsureValidId(int value, string parameterName)
    {
        if (value < 0)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                value,
                "The list id must not be negative; use null to clear the reference and 0 to say the paragraph is not numbered.");
        }
    }

    /// <summary>
    /// Throws when a list level is outside the range the schema allows.
    /// </summary>
    /// <param name="value">The level to check.</param>
    /// <param name="parameterName">The parameter the level came from.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="value"/> is negative or greater than <see cref="MaxLevel"/>.
    /// </exception>
    internal static void EnsureValidLevel(int value, string parameterName)
    {
        if (value is < 0 or > MaxLevel)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                value,
                $"The list level must be between 0 and {MaxLevel}, the range the schema allows.");
        }
    }

    /// <summary>
    /// Reads an integer attribute, reporting <see langword="null"/> when it is absent or its text
    /// cannot be parsed.
    /// </summary>
    /// <param name="value">The attribute value to read.</param>
    /// <returns>The parsed value, or <see langword="null"/>.</returns>
    internal static int? TryReadInt(Int32Value? value)
    {
        try
        {
            return value?.Value;
        }
        catch (FormatException)
        {
            return null;
        }
    }
}
