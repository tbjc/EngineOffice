using System;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A read-only view over a single level (<c>w:lvl</c>) of a list definition.
/// </summary>
/// <remarks>
/// <para>
/// A level belongs to an abstract numbering definition (<c>w:abstractNum</c>) inside the document's
/// numbering part (<c>word/numbering.xml</c>) and describes how one indentation level of a list is
/// numbered. This version reads levels but does not create or modify them beyond what
/// <see cref="WordDocument.AddList"/> produces.
/// </para>
/// <para>
/// Only the level's own numbering properties are exposed. The paragraph and run properties that a
/// level applies to its items (<c>w:pPr</c>/<c>w:rPr</c>) are preserved but not modelled.
/// </para>
/// <para>
/// This is a lightweight wrapper: it does not own the underlying content and must not be used after
/// the <see cref="WordDocument"/> that produced it has been disposed.
/// </para>
/// </remarks>
public sealed class WordListLevel
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private readonly W.Level _element;

    internal WordListLevel(W.Level element) => _element = element;

    /// <summary>
    /// Gets the level index (<c>w:lvl/@w:ilvl</c>), from zero for the outermost level to eight for
    /// the innermost one the schema allows.
    /// </summary>
    /// <exception cref="InvalidOfficeFileException">
    /// The level has no level index or the stored value cannot be parsed.
    /// </exception>
    public int Level
    {
        get
        {
            Int32Value? stored = _element.LevelIndex;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A list level stores no w:ilvl attribute, which the schema requires.");
            }

            return ReadStoredValue(() => stored.Value, "level index");
        }
    }

    /// <summary>
    /// Gets the value the level counts from (<c>w:start</c>), or <see langword="null"/> when the
    /// level does not set one.
    /// </summary>
    /// <exception cref="InvalidOfficeFileException">The level sets a start value without one, or a stored value cannot be parsed.</exception>
    public int? Start
    {
        get
        {
            W.StartNumberingValue? start = _element.StartNumberingValue;
            if (start is null)
            {
                return null;
            }

            Int32Value? stored = start.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A list level stores a w:start element without the required w:val attribute.");
            }

            return ReadStoredValue(() => stored.Value, "level start");
        }
    }

    /// <summary>
    /// Gets the numbering format of the level (<c>w:numFmt</c>), or <see langword="null"/> when the
    /// level does not declare one.
    /// </summary>
    /// <exception cref="InvalidOfficeFileException">The level sets a format without a value, or a stored value cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The level uses a number format that Office Open XML defines but this version does not model
    /// (for example <c>chicago</c> or <c>ideographDigital</c>).
    /// </exception>
    public WordListFormat? Format
    {
        get
        {
            W.NumberingFormat? format = _element.NumberingFormat;
            if (format is null)
            {
                return null;
            }

            EnumValue<W.NumberFormatValues>? stored = format.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A list level stores a w:numFmt element without the required w:val attribute.");
            }

            W.NumberFormatValues value = ReadStoredValue(() => stored.Value, "number format");
            return ToWordListFormat(value, format.GetAttribute("val", WordprocessingNamespace).Value);
        }
    }

    /// <summary>
    /// Gets the text pattern the level uses to build its numbering (<c>w:lvlText</c>), or
    /// <see langword="null"/> when the level does not set one.
    /// </summary>
    /// <remarks>
    /// The pattern uses placeholders such as <c>%1</c> for the current level and <c>%2</c> for a
    /// parent level; for example, <c>%1.</c> produces <c>1.</c>, <c>2.</c> and so on.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">The level sets a text pattern without a value.</exception>
    public string? Text
    {
        get
        {
            W.LevelText? text = _element.LevelText;
            if (text is null)
            {
                return null;
            }

            StringValue? stored = text.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A list level stores a w:lvlText element without the required w:val attribute.");
            }

            return stored.Value;
        }
    }

    /// <summary>
    /// Gets the justification of the level's numbering (<c>w:lvlJc</c>), or <see langword="null"/>
    /// when the level does not set one.
    /// </summary>
    /// <exception cref="InvalidOfficeFileException">The level sets a justification without a value, or a stored value cannot be parsed.</exception>
    public WordListJustification? Justification
    {
        get
        {
            W.LevelJustification? justification = _element.LevelJustification;
            if (justification is null)
            {
                return null;
            }

            EnumValue<W.LevelJustificationValues>? stored = justification.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A list level stores a w:lvlJc element without the required w:val attribute.");
            }

            return ToWordListJustification(ReadStoredValue(() => stored.Value, "level justification"));
        }
    }

    /// <summary>
    /// Finds the first level in <paramref name="abstractNumbering"/> whose index equals
    /// <paramref name="level"/>.
    /// </summary>
    internal static W.Level? Find(W.AbstractNum? abstractNumbering, int level)
    {
        if (abstractNumbering is null)
        {
            return null;
        }

        foreach (W.Level candidate in abstractNumbering.Elements<W.Level>())
        {
            if (ListReference.TryReadInt(candidate.LevelIndex) == level)
            {
                return candidate;
            }
        }

        return null;
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
                $"A list level stores a {propertyDescription} value that this document cannot read.",
                exception);
        }
    }

    private static WordListFormat ToWordListFormat(W.NumberFormatValues value, string? stored)
    {
        if (value == W.NumberFormatValues.Bullet)
        {
            return WordListFormat.Bullet;
        }

        if (value == W.NumberFormatValues.Decimal)
        {
            return WordListFormat.Decimal;
        }

        if (value == W.NumberFormatValues.LowerLetter)
        {
            return WordListFormat.LowerLetter;
        }

        if (value == W.NumberFormatValues.UpperLetter)
        {
            return WordListFormat.UpperLetter;
        }

        if (value == W.NumberFormatValues.LowerRoman)
        {
            return WordListFormat.LowerRoman;
        }

        if (value == W.NumberFormatValues.UpperRoman)
        {
            return WordListFormat.UpperRoman;
        }

        string described = string.IsNullOrEmpty(stored) ? value.ToString() : stored;
        throw new UnsupportedOfficeFeatureException(
            $"A list level uses the number format '{described}', a value that Office Open XML defines but this version does not model.");
    }

    private static WordListJustification ToWordListJustification(W.LevelJustificationValues value)
    {
        if (value == W.LevelJustificationValues.Left)
        {
            return WordListJustification.Left;
        }

        if (value == W.LevelJustificationValues.Center)
        {
            return WordListJustification.Center;
        }

        if (value == W.LevelJustificationValues.Right)
        {
            return WordListJustification.Right;
        }

        // ST_LevelJustification has exactly those three members, so this arm is defensive: an unknown
        // value is rejected while it is read, before it reaches this mapping.
        throw new UnsupportedOfficeFeatureException(
            $"A list level stores the justification '{value}', a value that Office Open XML defines but this version does not model.");
    }
}
