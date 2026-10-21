namespace OfficeEngine.Word;

/// <summary>
/// Validation shared by the members that take a style reference, such as
/// <see cref="Paragraph.StyleId"/> and <see cref="Run.StyleId"/>.
/// </summary>
/// <remarks>
/// <para>
/// A style reference is stored in an attribute (<c>w:pStyle/@w:val</c>, <c>w:rStyle/@w:val</c>), so
/// it must be storable as an attribute value and must satisfy the <c>ST_String</c> restriction the
/// schema places on it: it cannot be empty or longer than
/// <see cref="MaxLength"/> characters.
/// </para>
/// <para>
/// The style the reference names is deliberately not resolved here. Office Open XML has no
/// referential integrity for style references — a document whose paragraph points to a style that
/// does not exist is valid and Word renders it — so a dangling reference is not an error.
/// </para>
/// </remarks>
internal static class StyleReference
{
    /// <summary>
    /// The longest value the schema allows for a style identifier (<c>ST_String</c>). A longer value
    /// would make the produced package fail validation, so it is refused before it is written.
    /// </summary>
    internal const int MaxLength = 253;

    /// <summary>
    /// Throws when <paramref name="value"/> cannot be stored as a style reference.
    /// </summary>
    /// <param name="value">The style identifier to check.</param>
    /// <param name="parameterName">The parameter the identifier came from.</param>
    /// <exception cref="ArgumentException">
    /// <paramref name="value"/> is empty, is only whitespace, is longer than
    /// <see cref="MaxLength"/> characters, or carries a character that cannot be stored in an
    /// attribute.
    /// </exception>
    internal static void EnsureValid(string value, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException(
                "The style id must not be empty or whitespace; use null to clear the reference.",
                parameterName);
        }

        if (value.Length > MaxLength)
        {
            throw new ArgumentException(
                $"The style id must not be longer than {MaxLength} characters, the maximum the schema allows for a style reference.",
                parameterName);
        }

        PlainText.EnsureValidAttributeValue(value, parameterName);
    }
}
