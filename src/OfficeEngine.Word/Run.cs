using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over a single run of a paragraph.
/// </summary>
/// <remarks>
/// <para>
/// A run is the smallest unit that can carry character formatting. This handle is a lightweight
/// wrapper over the underlying content: reading or writing it mutates the document that produced
/// the run, and it does not own or dispose anything. Handles must not be used after the owning
/// <see cref="WordDocument"/> has been disposed.
/// </para>
/// <para>
/// Every formatting member is <em>tri-state</em>: <see langword="null"/> means the run does not
/// specify a value and therefore inherits it from styles or defaults, which is exactly what
/// Office Open XML does. Setting a value to <see langword="null"/> removes it again instead of
/// writing a different one, so no information is invented or silently inverted.
/// </para>
/// <para>
/// Properties that Office Open XML supports but this version does not model (for example
/// strike-through, highlighting, shadings or theme colours) are left untouched when a modelled
/// property is edited, so they are preserved across a save rather than dropped.
/// </para>
/// </remarks>
public sealed class Run
{
    /// <summary>
    /// Child order of the Office Open XML <c>CT_RPr</c> sequence. New elements are inserted
    /// relative to this order so that produced documents stay schema-valid.
    /// </summary>
    private static readonly string[] RunPropertyOrder =
    [
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
        "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
        "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
        "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
        "specVanish", "oMath",
    ];

    private readonly W.Run _element;

    internal Run(W.Run element) => _element = element;

    /// <summary>
    /// Gets or sets the text carried by this run.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The getter concatenates the run's own <c>w:t</c> text children; it does not interpret line
    /// breaks, tabs or drawings. Use <see cref="Content"/> for the faithful sequence of text, tabs
    /// and breaks.
    /// </para>
    /// <para>
    /// The setter replaces the run's content but keeps its character formatting (<c>w:rPr</c>) and
    /// removes any other content the run carried, because that content is being replaced by new
    /// text.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="value"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store. Add tabs and breaks with <see cref="AddTab"/> and <see cref="AddBreak"/>.
    /// </exception>
    public string Text
    {
        get
        {
            StringBuilder builder = new();
            foreach (W.Text text in _element.Elements<W.Text>())
            {
                builder.Append(text.Text);
            }

            return builder.ToString();
        }

        set
        {
            ArgumentNullException.ThrowIfNull(value);
            PlainText.EnsureValid(value, nameof(value));

            foreach (OpenXmlElement child in _element.ChildElements.ToArray())
            {
                if (child is not W.RunProperties)
                {
                    child.Remove();
                }
            }

            if (value.Length > 0)
            {
                _element.AppendChild(new W.Text(value) { Space = SpaceProcessingModeValues.Preserve });
            }
        }
    }

    /// <summary>
    /// Gets the content of this run that this version models, in document order.
    /// </summary>
    /// <remarks>
    /// Text appears as <see cref="RunText"/>, tab characters as <see cref="RunTab"/> and line
    /// breaks as <see cref="RunBreak"/>. Content that is preserved but not modelled is not
    /// returned here; see <see cref="RunContent"/> for the exact boundary.
    /// </remarks>
    /// <exception cref="UnsupportedOfficeFeatureException">A stored break uses a value this version does not model.</exception>
    /// <exception cref="InvalidOfficeFileException">A stored break cannot be read.</exception>
    public IReadOnlyList<RunContent> Content
    {
        get
        {
            List<RunContent> content = new();
            foreach (OpenXmlElement child in _element.ChildElements)
            {
                switch (child)
                {
                    case W.Text text:
                        content.Add(new RunText(text.Text));
                        break;
                    case W.TabChar:
                        content.Add(new RunTab());
                        break;
                    case W.Break lineBreak:
                        content.Add(new RunBreak(ReadBreakType(lineBreak)));
                        break;
                    case W.CarriageReturn:
                        // w:cr is a line break that carries no attributes.
                        content.Add(new RunBreak(BreakType.TextWrapping));
                        break;
                }
            }

            return content;
        }
    }

    /// <summary>
    /// Appends plain text to the end of this run.
    /// </summary>
    /// <param name="text">The text to append. An empty string is ignored.</param>
    /// <returns>This run, so that calls can be chained.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="text"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store. Add tabs and breaks with <see cref="AddTab"/> and <see cref="AddBreak"/>.
    /// </exception>
    public Run AddText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        PlainText.EnsureValid(text, nameof(text));

        if (text.Length > 0)
        {
            _element.AppendChild(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });
        }

        return this;
    }

    /// <summary>
    /// Appends a tab character to the end of this run.
    /// </summary>
    /// <returns>This run, so that calls can be chained.</returns>
    public Run AddTab()
    {
        _element.AppendChild(new W.TabChar());
        return this;
    }

    /// <summary>
    /// Appends an explicit line break to the end of this run.
    /// </summary>
    /// <param name="breakType">The kind of break to add.</param>
    /// <returns>This run, so that calls can be chained.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="breakType"/> is not a defined value.</exception>
    public Run AddBreak(BreakType breakType)
    {
        // Mapping first rejects an undefined value before the run is touched.
        W.BreakValues mapped = ToOoxmlBreakType(breakType);
        W.Break created = new();
        if (breakType != BreakType.TextWrapping)
        {
            // A text-wrapping break is expressed by omitting @w:type, matching what Word writes.
            created.Type = mapped;
        }

        _element.AppendChild(created);
        return this;
    }

    /// <summary>
    /// Gets or sets the identifier of the character style applied to the run
    /// (<c>w:rStyle/@w:val</c>).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see langword="null"/> means the run does not name a character style and inherits one, which
    /// is exactly what Office Open XML expresses by omitting <c>w:rStyle</c>.
    /// </para>
    /// <para>
    /// The reference is not resolved: a value that names a style the document does not define is
    /// stored as given. Office Open XML has no referential integrity here, and Word ignores the
    /// dangling reference. Use <see cref="WordDocument.FindStyle"/> to look the style up.
    /// </para>
    /// <para>
    /// Setting a value only touches this run's <c>w:rStyle</c>; the style definition itself is never
    /// modified. Any other content of <c>w:rStyle</c> is replaced.
    /// </para>
    /// </remarks>
    /// <value>The style id, or <see langword="null"/> when the run does not name a style.</value>
    /// <exception cref="ArgumentException">
    /// The assigned value is empty or whitespace, is longer than 253 characters, or contains a
    /// character that Office Open XML cannot store in an attribute.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">The run stores a style reference without a value.</exception>
    public string? StyleId
    {
        get
        {
            W.RunStyle? style = Properties?.GetFirstChild<W.RunStyle>();
            if (style is null)
            {
                return null;
            }

            StringValue? stored = style.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A run stores a w:rStyle element without the required w:val attribute.");
            }

            return stored.Value;
        }

        set
        {
            if (value is null)
            {
                Properties?.RemoveAllChildren<W.RunStyle>();
                PruneEmptyProperties();
                return;
            }

            // Validate before touching the tree so that a rejected value leaves no partial markup.
            StyleReference.EnsureValid(value, nameof(value));

            W.RunStyle configured = EnsureChild<W.RunStyle>(EnsureProperties());
            configured.Val = value;
        }
    }

    /// <summary>
    /// Gets or sets the run's font family, or <see langword="null"/> when it is inherited.
    /// </summary>
    /// <value>
    /// The Western (<c>w:ascii</c>) font name. When only <c>w:hAnsi</c> is present, its value is
    /// returned instead. East Asian and complex-script font attributes are not modelled and are
    /// preserved as they are.
    /// </value>
    /// <remarks>Setting <see langword="null"/> clears the Western and high-ANSI font attributes.</remarks>
    /// <exception cref="ArgumentException"><paramref name="value"/> is empty or whitespace.</exception>
    public string? FontFamily
    {
        get
        {
            W.RunFonts? fonts = Properties?.GetFirstChild<W.RunFonts>();
            if (fonts is null)
            {
                return null;
            }

            return fonts.Ascii?.Value ?? fonts.HighAnsi?.Value;
        }

        set
        {
            if (value is null)
            {
                W.RunFonts? fonts = Properties?.GetFirstChild<W.RunFonts>();
                if (fonts is not null)
                {
                    fonts.Ascii = null;
                    fonts.HighAnsi = null;
                    if (!fonts.HasAttributes && !fonts.HasChildren)
                    {
                        fonts.Remove();
                    }
                }

                PruneEmptyProperties();
                return;
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("The font family must not be empty or whitespace; use null to clear it.", nameof(value));
            }

            W.RunFonts configured = EnsureChild<W.RunFonts>(EnsureProperties());
            configured.Ascii = value;
            configured.HighAnsi = value;
        }
    }

    /// <summary>
    /// Gets or sets the font size in points, or <see langword="null"/> when it is inherited.
    /// </summary>
    /// <value>
    /// Office Open XML stores sizes in half-points, so sizes written by OfficeEngine are multiples
    /// of <c>0.5</c>. Reading returns the exact stored size.
    /// </value>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The assigned value is not a positive, finite number of points, or it is not representable in half-points.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored size is malformed.</exception>
    public double? FontSizePoints
    {
        get
        {
            W.FontSize? size = Properties?.GetFirstChild<W.FontSize>();
            if (size?.Val is null)
            {
                return null;
            }

            if (!int.TryParse(size.Val.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int halfPoints))
            {
                throw new InvalidOfficeFileException($"A run has a malformed font size '{size.Val.Value}'.");
            }

            return halfPoints / 2.0;
        }

        set
        {
            if (value is null)
            {
                Properties?.RemoveAllChildren<W.FontSize>();
                PruneEmptyProperties();
                return;
            }

            double points = value.Value;
            if (double.IsNaN(points) || double.IsInfinity(points) || points <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(value), value, "The font size must be a positive, finite number of points.");
            }

            double halfPoints = points * 2.0;
            if (halfPoints != Math.Truncate(halfPoints))
            {
                throw new ArgumentOutOfRangeException(
                    nameof(value),
                    value,
                    "Office Open XML stores font sizes in half-points, so the size must be a multiple of 0.5 points.");
            }

            W.FontSize configured = EnsureChild<W.FontSize>(EnsureProperties());
            configured.Val = ((long)Math.Truncate(halfPoints)).ToString(CultureInfo.InvariantCulture);
        }
    }

    /// <summary>
    /// Gets or sets whether the run is bold, or <see langword="null"/> when boldness is inherited.
    /// </summary>
    /// <remarks>
    /// <see langword="true"/> marks the run as bold and <see langword="false"/> explicitly turns
    /// boldness off, which is different from <see langword="null"/> (no property, inherit).
    /// </remarks>
    public bool? Bold
    {
        get
        {
            W.Bold? bold = Properties?.GetFirstChild<W.Bold>();
            if (bold is null)
            {
                return null;
            }

            OnOffValue? stored = bold.Val;
            if (stored is null)
            {
                return true;
            }

            return ReadStoredValue(() => stored.Value, "bold");
        }

        set
        {
            if (value is null)
            {
                Properties?.RemoveAllChildren<W.Bold>();
                PruneEmptyProperties();
                return;
            }

            W.Bold configured = EnsureChild<W.Bold>(EnsureProperties());
            if (value.Value)
            {
                // On is expressed by omitting @w:val, matching what Word itself writes.
                configured.Val = null;
            }
            else
            {
                configured.Val = false;
            }
        }
    }

    /// <summary>
    /// Gets or sets whether the run is italic, or <see langword="null"/> when italics are inherited.
    /// </summary>
    public bool? Italic
    {
        get
        {
            W.Italic? italic = Properties?.GetFirstChild<W.Italic>();
            if (italic is null)
            {
                return null;
            }

            OnOffValue? stored = italic.Val;
            if (stored is null)
            {
                return true;
            }

            return ReadStoredValue(() => stored.Value, "italic");
        }

        set
        {
            if (value is null)
            {
                Properties?.RemoveAllChildren<W.Italic>();
                PruneEmptyProperties();
                return;
            }

            W.Italic configured = EnsureChild<W.Italic>(EnsureProperties());
            if (value.Value)
            {
                configured.Val = null;
            }
            else
            {
                configured.Val = false;
            }
        }
    }

    /// <summary>
    /// Gets or sets the underline style, or <see langword="null"/> when the run has no underline
    /// property of its own.
    /// </summary>
    /// <remarks>
    /// <see cref="UnderlineStyle.None"/> explicitly turns underlining off, which is different from
    /// <see langword="null"/> (no property, inherit).
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">A stored style cannot be mapped.</exception>
    public UnderlineStyle? Underline
    {
        get
        {
            W.Underline? underline = Properties?.GetFirstChild<W.Underline>();
            EnumValue<W.UnderlineValues>? stored = underline?.Val;
            if (stored is null)
            {
                return null;
            }

            return FromOoxmlUnderline(ReadStoredValue(() => stored.Value, "underline"));
        }

        set
        {
            if (value is null)
            {
                Properties?.RemoveAllChildren<W.Underline>();
                PruneEmptyProperties();
                return;
            }

            // Validate before touching the tree so a rejected value leaves no partial markup behind.
            W.UnderlineValues ooxml = ToOoxmlUnderline(value.Value);
            W.Underline configured = EnsureChild<W.Underline>(EnsureProperties());
            configured.Val = ooxml;
        }
    }

    /// <summary>
    /// Gets or sets the explicit text colour, or <see langword="null"/> when the run has no explicit
    /// RGB colour.
    /// </summary>
    /// <value>
    /// The colour stored in <c>w:color/@w:val</c>. Theme colours and the <c>auto</c> value have no
    /// fixed RGB value and are therefore reported as <see langword="null"/>; their markup is left
    /// untouched so it survives a save.
    /// </value>
    /// <remarks>Setting <see langword="null"/> removes the <c>w:color</c> element, including any theme attributes.</remarks>
    public RgbColor? Color
    {
        get
        {
            W.Color? color = Properties?.GetFirstChild<W.Color>();
            if (color?.Val is null)
            {
                return null;
            }

            return RgbColor.TryParseHex(color.Val.Value, out RgbColor parsed) ? parsed : null;
        }

        set
        {
            if (value is null)
            {
                Properties?.RemoveAllChildren<W.Color>();
                PruneEmptyProperties();
                return;
            }

            W.Color configured = EnsureChild<W.Color>(EnsureProperties());
            configured.Val = value.Value.ToHex();
            configured.ThemeColor = null;
            configured.ThemeTint = null;
            configured.ThemeShade = null;
        }
    }

    /// <inheritdoc />
    public override string ToString() => Text;

    private W.RunProperties? Properties => _element.GetFirstChild<W.RunProperties>();

    /// <summary>Gets the run's property container, creating it as the first child when it is absent.</summary>
    private W.RunProperties EnsureProperties()
    {
        W.RunProperties? existing = Properties;
        if (existing is not null)
        {
            return existing;
        }

        W.RunProperties created = new();
        _element.InsertAt(created, 0);
        return created;
    }

    /// <summary>
    /// Returns the property of type <typeparamref name="T"/>, creating it in schema order when it is
    /// absent.
    /// </summary>
    private static T EnsureChild<T>(W.RunProperties properties)
        where T : OpenXmlElement, new()
    {
        T? existing = properties.GetFirstChild<T>();
        if (existing is not null)
        {
            return existing;
        }

        T created = new();
        InsertInOrder(properties, created);
        return created;
    }

    /// <summary>Inserts an element so that the <c>CT_RPr</c> child order required by the schema is kept.</summary>
    private static void InsertInOrder(W.RunProperties properties, OpenXmlElement element)
    {
        int rank = Rank(element.LocalName);
        foreach (OpenXmlElement child in properties.ChildElements)
        {
            if (Rank(child.LocalName) > rank)
            {
                properties.InsertBefore(element, child);
                return;
            }
        }

        properties.AppendChild(element);
    }

    /// <summary>
    /// Returns the schema position of a child element. Elements OfficeEngine does not know are
    /// treated as coming last, which keeps them after newly added modelled properties.
    /// </summary>
    private static int Rank(string localName)
    {
        int index = Array.IndexOf(RunPropertyOrder, localName);
        return index < 0 ? int.MaxValue : index;
    }

    /// <summary>Removes the property container when it no longer carries anything.</summary>
    private void PruneEmptyProperties()
    {
        W.RunProperties? properties = Properties;
        if (properties is not null && !properties.HasChildren && !properties.GetAttributes().Any())
        {
            properties.Remove();
        }
    }

    /// <summary>
    /// Reads a stored attribute value, turning the SDK's parse failure into a typed OfficeEngine error.
    /// </summary>
    /// <remarks>
    /// The Open XML SDK parses an on/off or enumeration attribute while reading it and throws a raw
    /// <see cref="FormatException"/> when the file carries a value it does not know. That behaviour
    /// was verified for every malformed on/off, <c>ST_Underline</c> and <c>ST_BrType</c> value
    /// OfficeEngine reads. OfficeEngine never lets such a framework exception escape: untrusted,
    /// malformed content is reported as <see cref="InvalidOfficeFileException"/> instead.
    /// </remarks>
    /// <param name="read">Reads the parsed value.</param>
    /// <param name="propertyDescription">Human-readable property name used in the error message.</param>
    /// <typeparam name="T">The parsed value type.</typeparam>
    /// <returns>The parsed value.</returns>
    private static T ReadStoredValue<T>(Func<T> read, string propertyDescription)
    {
        try
        {
            return read();
        }
        catch (FormatException exception)
        {
            throw new InvalidOfficeFileException(
                $"A run stores a {propertyDescription} value that this document cannot read.",
                exception);
        }
    }

    private static BreakType ReadBreakType(W.Break lineBreak)
    {
        EnumValue<W.BreakValues>? stored = lineBreak.Type;
        if (stored is null)
        {
            // @w:type defaults to textWrapping when it is omitted.
            return BreakType.TextWrapping;
        }

        W.BreakValues value = ReadStoredValue(() => stored.Value, "break type");
        if (value == W.BreakValues.Page)
        {
            return BreakType.Page;
        }

        if (value == W.BreakValues.Column)
        {
            return BreakType.Column;
        }

        if (value == W.BreakValues.TextWrapping)
        {
            return BreakType.TextWrapping;
        }

        throw new UnsupportedOfficeFeatureException(
            $"A run stores the break type '{value}', a value that Office Open XML defines but this version does not model.");
    }

    private static W.BreakValues ToOoxmlBreakType(BreakType value) => value switch
    {
        BreakType.TextWrapping => W.BreakValues.TextWrapping,
        BreakType.Page => W.BreakValues.Page,
        BreakType.Column => W.BreakValues.Column,
        _ => throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown break type."),
    };

    private static W.UnderlineValues ToOoxmlUnderline(UnderlineStyle value) => value switch
    {
        UnderlineStyle.None => W.UnderlineValues.None,
        UnderlineStyle.SingleLine => W.UnderlineValues.Single,
        UnderlineStyle.Words => W.UnderlineValues.Words,
        UnderlineStyle.DoubleLine => W.UnderlineValues.Double,
        UnderlineStyle.Thick => W.UnderlineValues.Thick,
        UnderlineStyle.Dotted => W.UnderlineValues.Dotted,
        UnderlineStyle.DottedHeavy => W.UnderlineValues.DottedHeavy,
        UnderlineStyle.Dash => W.UnderlineValues.Dash,
        UnderlineStyle.DashedHeavy => W.UnderlineValues.DashedHeavy,
        UnderlineStyle.DashLong => W.UnderlineValues.DashLong,
        UnderlineStyle.DashLongHeavy => W.UnderlineValues.DashLongHeavy,
        UnderlineStyle.DotDash => W.UnderlineValues.DotDash,
        UnderlineStyle.DashDotHeavy => W.UnderlineValues.DashDotHeavy,
        UnderlineStyle.DotDotDash => W.UnderlineValues.DotDotDash,
        UnderlineStyle.DashDotDotHeavy => W.UnderlineValues.DashDotDotHeavy,
        UnderlineStyle.Wave => W.UnderlineValues.Wave,
        UnderlineStyle.WavyHeavy => W.UnderlineValues.WavyHeavy,
        UnderlineStyle.WavyDouble => W.UnderlineValues.WavyDouble,
        _ => throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown underline style."),
    };

    /// <summary>
    /// Maps an Office Open XML underline value to its modelled style.
    /// </summary>
    /// <remarks>
    /// The SDK exposes <c>ST_Underline</c> as a value type with static properties rather than a
    /// C# enum, so its members cannot be used as <c>switch</c> patterns; they are matched through
    /// this table instead.
    /// </remarks>
    private static readonly Dictionary<W.UnderlineValues, UnderlineStyle> UnderlineFromOoxml =
        new Dictionary<W.UnderlineValues, UnderlineStyle>
        {
            [W.UnderlineValues.None] = UnderlineStyle.None,
            [W.UnderlineValues.Single] = UnderlineStyle.SingleLine,
            [W.UnderlineValues.Words] = UnderlineStyle.Words,
            [W.UnderlineValues.Double] = UnderlineStyle.DoubleLine,
            [W.UnderlineValues.Thick] = UnderlineStyle.Thick,
            [W.UnderlineValues.Dotted] = UnderlineStyle.Dotted,
            [W.UnderlineValues.DottedHeavy] = UnderlineStyle.DottedHeavy,
            [W.UnderlineValues.Dash] = UnderlineStyle.Dash,
            [W.UnderlineValues.DashedHeavy] = UnderlineStyle.DashedHeavy,
            [W.UnderlineValues.DashLong] = UnderlineStyle.DashLong,
            [W.UnderlineValues.DashLongHeavy] = UnderlineStyle.DashLongHeavy,
            [W.UnderlineValues.DotDash] = UnderlineStyle.DotDash,
            [W.UnderlineValues.DashDotHeavy] = UnderlineStyle.DashDotHeavy,
            [W.UnderlineValues.DotDotDash] = UnderlineStyle.DotDotDash,
            [W.UnderlineValues.DashDotDotHeavy] = UnderlineStyle.DashDotDotHeavy,
            [W.UnderlineValues.Wave] = UnderlineStyle.Wave,
            [W.UnderlineValues.WavyHeavy] = UnderlineStyle.WavyHeavy,
            [W.UnderlineValues.WavyDouble] = UnderlineStyle.WavyDouble,
        };

    private static UnderlineStyle FromOoxmlUnderline(W.UnderlineValues value)
        => UnderlineFromOoxml.TryGetValue(value, out UnderlineStyle style)
            ? style
            : throw new InvalidOfficeFileException($"A run uses an underline style '{value}' that this version does not model.");
}
