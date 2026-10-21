using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A read-only view over a single style definition (<c>w:style</c>) of a
/// <see cref="WordDocument"/>.
/// </summary>
/// <remarks>
/// <para>
/// A style definition lives in the document's style part (<c>word/styles.xml</c>) and can be based
/// on another style through <c>w:basedOn</c>, forming an inheritance chain. This version reads
/// styles and resolves that chain; it does not create or modify style definitions. A style that
/// this type does not model is still available from <see cref="WordDocument.Styles"/> and its
/// markup is preserved when the document is saved.
/// </para>
/// <para>
/// This is a lightweight wrapper: it does not own the underlying content and must not be used after
/// the <see cref="WordDocument"/> that produced it has been disposed.
/// </para>
/// </remarks>
public sealed class WordStyle
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private readonly W.Style _element;
    private readonly W.Styles? _root;

    internal WordStyle(W.Style element, W.Styles? root)
    {
        _element = element;
        _root = root;
    }

    /// <summary>Gets the style identifier (<c>w:styleId</c>), or <see langword="null"/> when the definition does not carry one.</summary>
    public string? Id => _element.StyleId?.Value;

    /// <summary>
    /// Gets the human-readable name of the style (<c>w:name</c>), or <see langword="null"/> when the
    /// style does not define one.
    /// </summary>
    /// <exception cref="InvalidOfficeFileException">The style names itself but the value is missing.</exception>
    public string? Name
    {
        get
        {
            W.StyleName? name = _element.GetFirstChild<W.StyleName>();
            if (name is null)
            {
                return null;
            }

            StringValue? stored = name.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A style stores a w:name element without the required w:val attribute.");
            }

            return stored.Value;
        }
    }

    /// <summary>
    /// Gets the kind of content the style applies to (<c>w:type</c>), or <see langword="null"/> when
    /// the style does not declare one.
    /// </summary>
    /// <exception cref="InvalidOfficeFileException">A stored type cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The style declares a type that Office Open XML defines but this version does not model.
    /// </exception>
    public WordStyleType? Type
    {
        get
        {
            EnumValue<W.StyleValues>? stored = _element.Type;
            return stored is null
                ? null
                : ToWordStyleType(ReadStoredValue(() => stored.Value, "type"));
        }
    }

    /// <summary>
    /// Gets whether the style is the default style of its type (<c>w:default</c>). Office Open XML
    /// omits the flag for a non-default style, so an absent attribute is reported as
    /// <see langword="false"/>.
    /// </summary>
    /// <exception cref="InvalidOfficeFileException">A stored flag cannot be parsed.</exception>
    public bool IsDefault
    {
        get
        {
            OnOffValue? stored = _element.Default;
            return stored is not null && ReadStoredValue(() => stored.Value, "default flag");
        }
    }

    /// <summary>
    /// Gets the identifier of the style this one is based on (<c>w:basedOn</c>), or
    /// <see langword="null"/> when it does not inherit from another style.
    /// </summary>
    /// <exception cref="InvalidOfficeFileException">A stored reference has no value.</exception>
    public string? BasedOnId => ReadBasedOnId(_element);

    /// <summary>
    /// Gets the formatting this style defines itself, without resolving inheritance.
    /// </summary>
    /// <remarks>
    /// Only the properties the style's own <c>w:rPr</c> and <c>w:pPr</c> declare are reported; the
    /// rest are <see langword="null"/>. Use <see cref="EffectiveFormatting"/> for the result of
    /// applying the <c>w:basedOn</c> chain.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">A stored property cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The style stores an alignment that Office Open XML defines but this version does not model.
    /// </exception>
    public WordStyleFormatting OwnFormatting => ReadFormatting(_element, default);

    /// <summary>
    /// Gets the formatting the style contributes after its <c>w:basedOn</c> chain has been resolved.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The chain is applied from the base style down to this one, so a property defined by a derived
    /// style overrides the value it inherits. A <c>w:basedOn</c> that points to a missing style is
    /// ignored, which is how Word treats it.
    /// </para>
    /// <para>
    /// Only the style chain is resolved. The document defaults (<c>w:docDefaults</c>), the default
    /// style of the type and the direct formatting of a paragraph or a run are outside this model
    /// and are not folded in.
    /// </para>
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">
    /// A stored property cannot be parsed, or the chain contains a <c>w:basedOn</c> cycle.
    /// </exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// A style in the chain stores an alignment that Office Open XML defines but this version does
    /// not model.
    /// </exception>
    public WordStyleFormatting EffectiveFormatting
    {
        get
        {
            WordStyleFormatting formatting = default;
            foreach (W.Style link in BuildInheritanceChain())
            {
                formatting = ReadFormatting(link, formatting);
            }

            return formatting;
        }
    }

    /// <inheritdoc />
    public override string ToString() => Id ?? "(style without an id)";

    /// <summary>
    /// Finds the first <c>w:style</c> in <paramref name="root"/> whose <c>w:styleId</c> equals
    /// <paramref name="styleId"/>, comparing with ordinal rules.
    /// </summary>
    internal static W.Style? Find(W.Styles? root, string styleId)
    {
        if (root is null)
        {
            return null;
        }

        foreach (W.Style candidate in root.Elements<W.Style>())
        {
            if (candidate.StyleId?.Value is { } id && string.Equals(id, styleId, StringComparison.Ordinal))
            {
                return candidate;
            }
        }

        return null;
    }

    private List<W.Style> BuildInheritanceChain()
    {
        List<W.Style> chain = [_element];
        W.Style current = _element;
        while (true)
        {
            string? basedOn = ReadBasedOnId(current);
            if (basedOn is null)
            {
                break;
            }

            W.Style? baseStyle = Find(_root, basedOn);
            if (baseStyle is null)
            {
                break;
            }

            foreach (W.Style link in chain)
            {
                if (ReferenceEquals(link, baseStyle))
                {
                    throw new InvalidOfficeFileException(
                        $"A style inherits from '{basedOn}' through a w:basedOn cycle that this document cannot resolve.");
                }
            }

            chain.Add(baseStyle);
            current = baseStyle;
        }

        chain.Reverse();
        return chain;
    }

    private static WordStyleFormatting ReadFormatting(W.Style style, WordStyleFormatting inherited)
    {
        W.StyleRunProperties? run = style.GetFirstChild<W.StyleRunProperties>();
        W.StyleParagraphProperties? paragraph = style.GetFirstChild<W.StyleParagraphProperties>();
        W.Indentation? indentation = paragraph?.Indentation;
        W.SpacingBetweenLines? spacing = paragraph?.SpacingBetweenLines;

        return new WordStyleFormatting(
            fontFamily: ReadFontFamily(run) ?? inherited.FontFamily,
            fontSizePoints: ReadFontSize(run) ?? inherited.FontSizePoints,
            bold: ReadOnOff(run?.Bold, "bold") ?? inherited.Bold,
            italic: ReadOnOff(run?.Italic, "italic") ?? inherited.Italic,
            color: run?.Color is { } color ? ReadColor(color) : inherited.Color,
            alignment: ReadAlignment(paragraph) ?? inherited.Alignment,
            leftIndentPoints: ReadTwips(indentation?.Left, "left indent") ?? inherited.LeftIndentPoints,
            rightIndentPoints: ReadTwips(indentation?.Right, "right indent") ?? inherited.RightIndentPoints,
            firstLineIndentPoints: ReadTwips(indentation?.FirstLine, "first-line indent") ?? inherited.FirstLineIndentPoints,
            hangingIndentPoints: ReadTwips(indentation?.Hanging, "hanging indent") ?? inherited.HangingIndentPoints,
            spacingBeforePoints: ReadTwips(spacing?.Before, "spacing before") ?? inherited.SpacingBeforePoints,
            spacingAfterPoints: ReadTwips(spacing?.After, "spacing after") ?? inherited.SpacingAfterPoints);
    }

    private static string? ReadBasedOnId(W.Style style)
    {
        W.BasedOn? basedOn = style.GetFirstChild<W.BasedOn>();
        if (basedOn is null)
        {
            return null;
        }

        StringValue? val = basedOn.Val;
        if (val is null)
        {
            throw new InvalidOfficeFileException("A style stores a w:basedOn reference without the required w:val attribute.");
        }

        return val.Value;
    }

    private static string? ReadFontFamily(W.StyleRunProperties? run)
    {
        W.RunFonts? fonts = run?.RunFonts;
        return fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
    }

    private static double? ReadFontSize(W.StyleRunProperties? run)
    {
        StringValue? size = run?.FontSize?.Val;
        if (size is null)
        {
            return null;
        }

        if (!int.TryParse(size.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int halfPoints))
        {
            throw new InvalidOfficeFileException($"A style stores a malformed font size '{size.Value}'.");
        }

        return halfPoints / 2.0;
    }

    private static bool? ReadOnOff(W.OnOffType? element, string propertyDescription)
    {
        if (element is null)
        {
            return null;
        }

        OnOffValue? stored = element.Val;
        if (stored is null)
        {
            // An on/off element without w:val, such as <w:b/>, means "on".
            return true;
        }

        return ReadStoredValue(() => stored.Value, propertyDescription);
    }

    private static RgbColor? ReadColor(W.Color color)
    {
        StringValue? value = color.Val;
        return value is not null && RgbColor.TryParseHex(value.Value, out RgbColor parsed) ? parsed : null;
    }

    private static ParagraphAlignment? ReadAlignment(W.StyleParagraphProperties? paragraph)
    {
        EnumValue<W.JustificationValues>? stored = paragraph?.Justification?.Val;
        return stored is null ? null : ToParagraphAlignment(ReadStoredValue(() => stored.Value, "alignment"));
    }

    private static double? ReadTwips(StringValue? stored, string propertyDescription)
    {
        if (stored is null)
        {
            return null;
        }

        if (!int.TryParse(stored.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips))
        {
            throw new InvalidOfficeFileException($"A style stores a {propertyDescription} of '{stored.Value}' that this document cannot read.");
        }

        return twips / 20.0;
    }

    private static ParagraphAlignment ToParagraphAlignment(W.JustificationValues value)
    {
        if (value == W.JustificationValues.Left)
        {
            return ParagraphAlignment.Left;
        }

        if (value == W.JustificationValues.Center)
        {
            return ParagraphAlignment.Center;
        }

        if (value == W.JustificationValues.Right)
        {
            return ParagraphAlignment.Right;
        }

        if (value == W.JustificationValues.Both)
        {
            return ParagraphAlignment.Both;
        }

        if (value == W.JustificationValues.Distribute)
        {
            return ParagraphAlignment.Justify;
        }

        throw new UnsupportedOfficeFeatureException(
            $"A style stores the alignment '{value}', a value that Office Open XML defines but this version does not model.");
    }

    private static WordStyleType ToWordStyleType(W.StyleValues value)
    {
        if (value == W.StyleValues.Paragraph)
        {
            return WordStyleType.Paragraph;
        }

        if (value == W.StyleValues.Character)
        {
            return WordStyleType.Character;
        }

        if (value == W.StyleValues.Table)
        {
            return WordStyleType.Table;
        }

        if (value == W.StyleValues.Numbering)
        {
            return WordStyleType.Numbering;
        }

        // ST_Style has exactly those four members, so this arm is defensive: an unknown value is
        // rejected while it is read, before it reaches this mapping.
        throw new UnsupportedOfficeFeatureException(
            $"A style has the type '{value}', a value that Office Open XML defines but this version does not model.");
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
                $"A style stores a {propertyDescription} value that this document cannot read.",
                exception);
        }
    }
}
