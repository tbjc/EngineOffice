using System;
using System.Xml;

namespace OfficeEngine.Word;

/// <summary>
/// Validation shared by the entry points that take plain text.
/// </summary>
/// <remarks>
/// <para>
/// Office Open XML does not express a tab or a line break as text: they are dedicated elements
/// (<c>w:tab</c>, <c>w:br</c>) inside a run. A tab or newline character stored in <c>w:t</c> is
/// schema-valid but is not what the caller asked for, so OfficeEngine refuses it instead of writing
/// something that renders differently. The dedicated members (<see cref="Run.AddTab"/>,
/// <see cref="Run.AddBreak"/>) build that content explicitly.
/// </para>
/// <para>
/// The same check refuses characters that Office Open XML cannot store at all, such as a null
/// character or a lone surrogate, which would otherwise surface as a framework exception when the
/// document is saved.
/// </para>
/// </remarks>
internal static class PlainText
{
    /// <summary>
    /// Throws when <paramref name="value"/> cannot be stored as plain Office Open XML text.
    /// </summary>
    /// <param name="value">The text to check.</param>
    /// <param name="parameterName">Name of the parameter being validated.</param>
    /// <exception cref="ArgumentException">
    /// <paramref name="value"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store.
    /// </exception>
    internal static void EnsureValid(string value, string parameterName)
    {
        for (int index = 0; index < value.Length; index++)
        {
            char character = value[index];
            if (character is '\t' or '\n' or '\r')
            {
                throw new ArgumentException(
                    "The text must not contain tab or line-break characters; add them explicitly with Run.AddTab or Run.AddBreak.",
                    parameterName);
            }

            if (char.IsHighSurrogate(character))
            {
                if (index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1]))
                {
                    throw UnsupportedCharacter(character, parameterName);
                }

                // A well-formed surrogate pair is a single code point in the XML character range.
                index++;
                continue;
            }

            if (char.IsLowSurrogate(character) || !XmlConvert.IsXmlChar(character))
            {
                throw UnsupportedCharacter(character, parameterName);
            }
        }
    }

    /// <summary>
    /// Throws when <paramref name="value"/> cannot be stored as an Office Open XML attribute value.
    /// </summary>
    /// <remarks>
    /// An attribute value carries the same restriction as plain text (no null character or lone
    /// surrogate), plus the one XML itself imposes on attributes: a tab or a line break is folded
    /// into a space while writing, so a value carrying one would not read back as it was set.
    /// </remarks>
    /// <param name="value">The value to check.</param>
    /// <param name="parameterName">The parameter the value came from.</param>
    /// <exception cref="ArgumentException">
    /// <paramref name="value"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store.
    /// </exception>
    internal static void EnsureValidAttributeValue(string value, string parameterName)
    {
        for (int index = 0; index < value.Length; index++)
        {
            char character = value[index];
            if (character is '\t' or '\n' or '\r')
            {
                throw new ArgumentException(
                    "The value must not contain tab or line-break characters because XML folds them into spaces while writing.",
                    parameterName);
            }

            if (char.IsHighSurrogate(character))
            {
                if (index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1]))
                {
                    throw new ArgumentException(
                        $"The value contains the character U+{(int)character:X4}, which Office Open XML cannot store.",
                        parameterName);
                }

                index++;
                continue;
            }

            if (char.IsLowSurrogate(character) || !XmlConvert.IsXmlChar(character))
            {
                throw new ArgumentException(
                    $"The value contains the character U+{(int)character:X4}, which Office Open XML cannot store.",
                    parameterName);
            }
        }
    }

    private static ArgumentException UnsupportedCharacter(char character, string parameterName)
        => new(
            $"The text contains the character U+{(int)character:X4}, which Office Open XML cannot store.",
            parameterName);
}
