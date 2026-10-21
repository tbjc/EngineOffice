using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over a single paragraph of a <see cref="WordDocument"/>.
/// </summary>
/// <remarks>
/// <para>
/// This is a lightweight wrapper: it does not own the underlying content. Reading or writing
/// text mutates the document that produced the paragraph, and handles must not be used after
/// that document has been disposed.
/// </para>
/// <para>
/// Character format is modelled per <see cref="Run"/>; paragraph format (alignment, indentation
/// and spacing) is modelled here. A value that Office Open XML defines but this version does not
/// model is never guessed: reading it throws, and the original markup is preserved.
/// </para>
/// </remarks>
public sealed class Paragraph
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Child order required by CT_PPr. The names are the XML local names (not the SDK property names)
    // and the sequence is the one in which the Open XML SDK generates the children of a
    // ParagraphProperties element, verified with a reflection probe over DocumentFormat.OpenXml 3.5.1.
    // Properties are inserted following this table instead of being appended, so the package stays
    // valid even when a paragraph already carries markup that this version does not model (for
    // example a style, a bidirectional flag or a border).
    private static readonly string[] ParagraphPropertyOrder =
    [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
        "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
        "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
        "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
        "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
    ];

    private static readonly string[] NumberingPropertyOrder =
    [
        "ilvl", "numId", "numberingChange", "ins",
    ];

    private readonly W.Paragraph _element;

    internal Paragraph(W.Paragraph element) => _element = element;

    internal W.Paragraph Element => _element;

    /// <summary>
    /// Finds the occurrences of a text inside this paragraph.
    /// </summary>
    /// <param name="text">The text to look for. It cannot be empty.</param>
    /// <param name="matchCase">Whether the comparison distinguishes upper and lower case.</param>
    /// <returns>The occurrences, in order of appearance.</returns>
    /// <remarks>
    /// The searched text is the one of the direct runs of the paragraph and of the runs of its
    /// direct hyperlinks. The search never crosses the boundary of the paragraph.
    /// </remarks>
    /// <exception cref="ArgumentException"><paramref name="text"/> is empty.</exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public IReadOnlyList<WordTextMatch> FindText(string text, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrEmpty(text);
        return WordTextSearch.Find(this, text, matchCase);
    }

    /// <summary>Replaces the occurrences of a text inside this paragraph.</summary>
    /// <param name="oldText">The text to replace. It cannot be empty.</param>
    /// <param name="newText">The text that replaces it. An empty text removes the match.</param>
    /// <param name="matchCase">Whether the comparison distinguishes upper and lower case.</param>
    /// <returns>The number of replacements.</returns>
    /// <remarks>
    /// The replaced text keeps the character formatting of the <em>first</em> run of the match, the
    /// same rule that Word applies. The rest of the text of the match is removed from the runs that
    /// carried it, and the text before and after the match stays where it was.
    /// </remarks>
    /// <exception cref="ArgumentException"><paramref name="oldText"/> is empty.</exception>
    /// <exception cref="ArgumentNullException"><paramref name="newText"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="newText"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public int ReplaceText(string oldText, string newText, bool matchCase = false)
        => WordTextSearch.Replace(this, oldText, newText, matchCase);

    private W.ParagraphProperties? Properties
        => _element.GetFirstChild<W.ParagraphProperties>();

    /// <summary>
    /// Appends a field to the end of this paragraph.
    /// </summary>
    /// <param name="instruction">
    /// The instruction of the field, for example <c>PAGE</c> or <c>NUMPAGES</c>. It cannot be empty.
    /// </param>
    /// <param name="cachedResult">
    /// The result stored with the field. It is what a reader shows until the engine that opens the
    /// document recalculates the field; when it is not given, the field is written without result.
    /// </param>
    /// <returns>The field that was written.</returns>
    /// <remarks>
    /// <para>
    /// The field is written as the sequence of runs that Word uses: the character that opens the
    /// field, the instruction, the character that separates the instruction from the result, the
    /// result and the character that closes the field.
    /// </para>
    /// <para>
    /// This library never calculates a field. The result it stores is the one that was given, and
    /// Word recomputes it when the document is opened or printed.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentException"><paramref name="instruction"/> is empty.</exception>
    /// <exception cref="ArgumentException">
    /// The instruction or the result contains a tab, a line break or a character that Office Open XML
    /// cannot store.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public WordField AddField(string instruction, string? cachedResult = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(instruction);
        PlainText.EnsureValid(instruction, nameof(instruction));
        if (cachedResult is not null)
        {
            PlainText.EnsureValid(cachedResult, nameof(cachedResult));
        }

        string code = " " + instruction.Trim() + " ";
        _element.AppendChild(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }));
        _element.AppendChild(new W.Run(
            new W.FieldCode(code) { Space = SpaceProcessingModeValues.Preserve }));
        if (cachedResult is not null)
        {
            _element.AppendChild(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }));
            _element.AppendChild(new W.Run(
                new W.Text(cachedResult) { Space = SpaceProcessingModeValues.Preserve }));
        }

        _element.AppendChild(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        return new WordField(instruction.Trim(), cachedResult);
    }

    /// <summary>
    /// Gets or sets the visible text of the paragraph.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The getter concatenates the text of every text element in the paragraph, including runs
    /// nested in hyperlinks and other structures. Tabs, line breaks and other non-text content
    /// contribute nothing to the string; use <see cref="Run.Content"/> to read them.
    /// </para>
    /// <para>
    /// The setter replaces the paragraph's content with a single run carrying the new text.
    /// Paragraph properties (<c>w:pPr</c>), such as alignment, indentation or spacing, are
    /// preserved; other structures that carried the old text, such as hyperlinks or bookmarks, are
    /// removed together with it.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="value"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store. Add tabs and breaks with <see cref="Run.AddTab"/> and <see cref="Run.AddBreak"/>.
    /// </exception>
    public string Text
    {
        get
        {
            StringBuilder builder = new();
            foreach (W.Text text in _element.Descendants<W.Text>())
            {
                builder.Append(text.Text);
            }

            return builder.ToString();
        }

        set
        {
            ArgumentNullException.ThrowIfNull(value);
            PlainText.EnsureValid(value, nameof(value));

            RemoveContentKeepingParagraphProperties();
            if (value.Length > 0)
            {
                _element.AppendChild(CreateRun(value));
            }
        }
    }

    /// <summary>
    /// Gets a live view over every run that is a direct child of this paragraph.
    /// </summary>
    /// <remarks>
    /// The list reflects the current state of the paragraph and is rebuilt on each access, so it
    /// always reflects the current content. Runs nested inside structures that are not modelled in
    /// this version (for example content controls) are not returned here, although their text is
    /// still included by <see cref="Text"/>. Runs inside a modelled external hyperlink are
    /// available through <see cref="Hyperlink.Runs"/>.
    /// </remarks>
    public IReadOnlyList<Run> Runs
    {
        get
        {
            List<Run> runs = new();
            foreach (W.Run element in _element.Elements<W.Run>())
            {
                runs.Add(new Run(element));
            }

            return runs;
        }
    }

    /// <summary>
    /// Gets a live snapshot of the embedded inline images in this paragraph's direct runs, in order.
    /// </summary>
    /// <remarks>
    /// Only images whose drawing is a direct child of a direct run are returned. Images nested in
    /// hyperlink runs, content controls or other structures this version does not model, floating
    /// images (<c>wp:anchor</c>), linked images (<c>r:link</c>) and non-picture graphics such as
    /// charts are preserved but not returned here.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">
    /// A stored inline image is malformed, or the relationship it references is missing or is not an
    /// image part.
    /// </exception>
    public IReadOnlyList<Image> Images
    {
        get
        {
            List<Image> images = new();
            foreach (W.Run element in _element.Elements<W.Run>())
            {
                Image.AddRunImages(element, images);
            }

            return images;
        }
    }

    /// <summary>
    /// Gets a live snapshot of the external hyperlinks that are direct children of this paragraph,
    /// in order.
    /// </summary>
    /// <remarks>
    /// Only links whose <c>w:hyperlink</c> carries an <c>r:id</c> that resolves to an external
    /// hyperlink relationship are returned. Internal links to a bookmark (<c>w:anchor</c>), links
    /// without a relationship and links nested in content controls or other structures this version
    /// does not model are preserved but not returned here.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">
    /// A link references a relationship that is missing, is not a hyperlink relationship, or is not
    /// external.
    /// </exception>
    public IReadOnlyList<Hyperlink> Hyperlinks
    {
        get
        {
            List<Hyperlink> hyperlinks = new();
            foreach (W.Hyperlink element in _element.Elements<W.Hyperlink>())
            {
                if (element.Id is null)
                {
                    // An internal bookmark link (w:anchor) or a link without a target is valid but
                    // outside what this version models.
                    continue;
                }

                Hyperlink.GetRelationship(element);
                hyperlinks.Add(new Hyperlink(element));
            }

            return hyperlinks;
        }
    }

    /// <summary>
    /// Adds a new run with the specified text to the end of this paragraph.
    /// </summary>
    /// <param name="text">The text content of the run. May be empty.</param>
    /// <returns>A live view over the new run.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="text"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store. Add tabs and breaks with <see cref="Run.AddTab"/> and <see cref="Run.AddBreak"/>.
    /// </exception>
    public Run AddRun(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        PlainText.EnsureValid(text, nameof(text));

        W.Run runElement = text.Length > 0 ? CreateRun(text) : new W.Run();
        _element.AppendChild(runElement);
        return new Run(runElement);
    }

    /// <summary>
    /// Adds an external hyperlink at the end of this paragraph.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The destination is stored as a hyperlink relationship owned by the part that contains the
    /// paragraph (the main document, a header or a footer). Each call creates its own relationship;
    /// links with the same destination are not deduplicated. The link is appended as a new
    /// <c>w:hyperlink</c>, so the paragraph's existing runs and text are not changed.
    /// </para>
    /// <para>
    /// No character style or direct formatting is applied to the link text. Use
    /// <see cref="Hyperlink.Runs"/> to format it, or read and edit the text and destination through
    /// <see cref="Hyperlink"/>. Internal links to a bookmark are not created by this version.
    /// </para>
    /// </remarks>
    /// <param name="text">The visible text of the link. May be empty.</param>
    /// <param name="destination">The external destination. It is not opened or validated by scheme.</param>
    /// <returns>A live view over the new hyperlink.</returns>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="text"/> or <paramref name="destination"/> is <see langword="null"/>.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="text"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The owning document has been disposed.</exception>
    public Hyperlink AddHyperlink(string text, Uri destination)
    {
        ArgumentNullException.ThrowIfNull(text);
        PlainText.EnsureValid(text, nameof(text));
        ArgumentNullException.ThrowIfNull(destination);

        return Hyperlink.Add(_element, text, destination);
    }

    /// <summary>
    /// Adds an inline image from a file at the end of this paragraph.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The format is inferred from the file extension. The image is copied into the package as a new
    /// image part with its own relationship; adding the same file twice creates two parts, because
    /// this version does not deduplicate image data.
    /// </para>
    /// <para>
    /// The image is appended in a new run. The paragraph's existing text and runs are not changed.
    /// Only inline images are supported: floating images, linked images and alt text are outside this
    /// version (see <see cref="Image"/>).
    /// </para>
    /// </remarks>
    /// <param name="path">Path of the image file. The extension must be supported.</param>
    /// <param name="widthPoints">Displayed width in points. Must be positive and finite.</param>
    /// <param name="heightPoints">Displayed height in points. Must be positive and finite.</param>
    /// <returns>A live view over the added image.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is empty, its extension is not supported, or a dimension is not a
    /// positive, finite number of points within the range Office Open XML can store.
    /// </exception>
    /// <exception cref="IOException">The file cannot be read.</exception>
    /// <exception cref="ObjectDisposedException">The owning document has been disposed.</exception>
    public Image AddImage(string path, double widthPoints, double heightPoints)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);

        ImageFormat format = Image.FromPath(path);
        long widthEmus = Image.ToEmus(widthPoints, nameof(widthPoints));
        long heightEmus = Image.ToEmus(heightPoints, nameof(heightPoints));

        using FileStream stream = File.OpenRead(path);
        return AddImageCore(stream, format, widthEmus, heightEmus, Path.GetFileName(path));
    }

    /// <summary>
    /// Adds an inline image from a stream at the end of this paragraph.
    /// </summary>
    /// <remarks>
    /// The stream is read from its current position to its end; it is not disposed. See
    /// <see cref="AddImage(string, double, double)"/> for the package and formatting behaviour.
    /// </remarks>
    /// <param name="stream">Readable stream with the image bytes.</param>
    /// <param name="format">The image format, which fixes the part's content type.</param>
    /// <param name="widthPoints">Displayed width in points. Must be positive and finite.</param>
    /// <param name="heightPoints">Displayed height in points. Must be positive and finite.</param>
    /// <returns>A live view over the added image.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="stream"/> is not readable, or a dimension is not a positive, finite number of
    /// points within the range Office Open XML can store.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is not a defined value.</exception>
    /// <exception cref="ObjectDisposedException">The owning document has been disposed.</exception>
    public Image AddImage(Stream stream, ImageFormat format, double widthPoints, double heightPoints)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead)
        {
            throw new ArgumentException("The stream must be readable.", nameof(stream));
        }

        Image.EnsureDefined(format, nameof(format));
        long widthEmus = Image.ToEmus(widthPoints, nameof(widthPoints));
        long heightEmus = Image.ToEmus(heightPoints, nameof(heightPoints));

        return AddImageCore(stream, format, widthEmus, heightEmus, Image.NameFor(format));
    }

    private Image AddImageCore(
        Stream stream,
        ImageFormat format,
        long widthEmus,
        long heightEmus,
        string name)
    {
        OpenXmlPart owner = Image.ResolvePart(_element);
        Image image = Image.CreateInline(owner, stream, format, widthEmus, heightEmus, name);
        _element.AppendChild(new W.Run(image.Drawing));
        return image;
    }

    /// <summary>
    /// Appends text to the paragraph as a new run.
    /// </summary>
    /// <param name="text">The text to append. An empty string is ignored.</param>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="text"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store. Add tabs and breaks with <see cref="Run.AddTab"/> and <see cref="Run.AddBreak"/>.
    /// </exception>
    public void AppendText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        PlainText.EnsureValid(text, nameof(text));

        if (text.Length > 0)
        {
            _element.AppendChild(CreateRun(text));
        }
    }

    /// <summary>
    /// Gets or sets the identifier of the paragraph style applied to the paragraph
    /// (<c>w:pStyle/@w:val</c>).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see langword="null"/> means the paragraph does not name a style and inherits one from the
    /// document defaults, which is exactly what Office Open XML expresses by omitting <c>w:pStyle</c>.
    /// </para>
    /// <para>
    /// The reference is not resolved: a value that names a style the document does not define is
    /// stored as given. Office Open XML has no referential integrity here, and Word ignores the
    /// dangling reference. Use <see cref="WordDocument.FindStyle"/> to look the style up.
    /// </para>
    /// <para>
    /// Setting a value only touches this paragraph's <c>w:pStyle</c>; the style definition itself is
    /// never modified. Any other content of <c>w:pStyle</c> is replaced.
    /// </para>
    /// </remarks>
    /// <value>The style id, or <see langword="null"/> when the paragraph does not name a style.</value>
    /// <exception cref="ArgumentException">
    /// The assigned value is empty or whitespace, is longer than 253 characters, or contains a
    /// character that Office Open XML cannot store in an attribute.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">The paragraph stores a style reference without a value.</exception>
    public string? StyleId
    {
        get
        {
            W.ParagraphStyleId? style = Properties?.GetFirstChild<W.ParagraphStyleId>();
            if (style is null)
            {
                return null;
            }

            StringValue? stored = style.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A paragraph stores a w:pStyle element without the required w:val attribute.");
            }

            return stored.Value;
        }

        set
        {
            if (value is null)
            {
                W.ParagraphProperties? properties = Properties;
                if (properties is not null)
                {
                    properties.RemoveAllChildren<W.ParagraphStyleId>();
                    PruneParagraphPropertiesIfEmpty();
                }

                return;
            }

            // Validate before touching the tree so that a rejected value leaves no partial markup.
            StyleReference.EnsureValid(value, nameof(value));

            W.ParagraphProperties owned = EnsureParagraphProperties();
            owned.RemoveAllChildren<W.ParagraphStyleId>();
            InsertProperty(owned, new W.ParagraphStyleId { Val = value });
        }
    }

    /// <summary>
    /// Gets or sets the identifier of the list this paragraph belongs to (<c>w:numPr/w:numId</c>).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see langword="null"/> means the paragraph does not name a list, which is how Office Open XML
    /// expresses "inherit whatever the style says". Assigning <c>0</c> is different: the schema
    /// reserves that value to say the paragraph is explicitly not numbered, overriding a list it
    /// would otherwise inherit.
    /// </para>
    /// <para>
    /// The reference is not resolved: a value that names a list the document does not define is
    /// stored as given. Office Open XML has no referential integrity here, and Word renders a
    /// paragraph whose list is missing as if it were not numbered. Use
    /// <see cref="WordDocument.FindList"/> to look the list up.
    /// </para>
    /// <para>
    /// Setting a value only touches this paragraph's <c>w:numPr</c>; the list definition itself is
    /// never modified. Use <see cref="ApplyList"/> to point a paragraph at a list without setting
    /// the identifier and the level separately.
    /// </para>
    /// </remarks>
    /// <value>The list identifier, or <see langword="null"/> when the paragraph does not name a list.</value>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is negative.</exception>
    /// <exception cref="InvalidOfficeFileException">
    /// The paragraph stores a list reference without a value, or a stored value cannot be parsed.
    /// </exception>
    public int? ListId
    {
        get
        {
            W.NumberingId? numbering = Properties?.GetFirstChild<W.NumberingProperties>()?.GetFirstChild<W.NumberingId>();
            if (numbering is null)
            {
                return null;
            }

            Int32Value? stored = numbering.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A paragraph stores a w:numId element without the required w:val attribute.");
            }

            return ReadStoredValue(() => stored.Value, "list id");
        }

        set
        {
            if (value is null)
            {
                RemoveNumberingChild<W.NumberingId>();
                return;
            }

            // Validate before touching the tree so that a rejected value leaves no partial markup.
            ListReference.EnsureValidId(value.Value, nameof(value));

            W.ParagraphProperties owned = EnsureParagraphProperties();
            W.NumberingProperties numbering = EnsureNumberingProperties(owned);
            numbering.RemoveAllChildren<W.NumberingId>();
            InsertNumberingProperty(numbering, new W.NumberingId { Val = value.Value });
        }
    }

    /// <summary>
    /// Gets or sets the level of the list this paragraph is at (<c>w:numPr/w:ilvl</c>), from zero for
    /// the outermost level to eight for the innermost one the schema allows.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see langword="null"/> means the paragraph does not name a level, which is how Office Open XML
    /// expresses "inherit the level": Word then uses the level of the style, or the first level of
    /// the list.
    /// </para>
    /// <para>
    /// Setting a value only touches this paragraph's <c>w:numPr</c>. Reading reports whatever the
    /// document stores, including a level outside the range this version writes.
    /// </para>
    /// </remarks>
    /// <value>The list level, or <see langword="null"/> when the paragraph does not name one.</value>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is not between 0 and 8.</exception>
    /// <exception cref="InvalidOfficeFileException">
    /// The paragraph stores a level without a value, or a stored value cannot be parsed.
    /// </exception>
    public int? ListLevel
    {
        get
        {
            W.NumberingLevelReference? level = Properties?.GetFirstChild<W.NumberingProperties>()?.GetFirstChild<W.NumberingLevelReference>();
            if (level is null)
            {
                return null;
            }

            Int32Value? stored = level.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A paragraph stores a w:ilvl element without the required w:val attribute.");
            }

            return ReadStoredValue(() => stored.Value, "list level");
        }

        set
        {
            if (value is null)
            {
                RemoveNumberingChild<W.NumberingLevelReference>();
                return;
            }

            // Validate before touching the tree so that a rejected value leaves no partial markup.
            ListReference.EnsureValidLevel(value.Value, nameof(value));

            W.ParagraphProperties owned = EnsureParagraphProperties();
            W.NumberingProperties numbering = EnsureNumberingProperties(owned);
            numbering.RemoveAllChildren<W.NumberingLevelReference>();
            InsertNumberingProperty(numbering, new W.NumberingLevelReference { Val = value.Value });
        }
    }

    /// <summary>
    /// Applies a list to the paragraph, placing it at the given level.
    /// </summary>
    /// <remarks>
    /// This writes the list's identifier to <see cref="ListId"/> and the level to
    /// <see cref="ListLevel"/>; it does not copy the list's formatting onto the paragraph, which
    /// Office Open XML resolves from the list definition when the document is rendered.
    /// </remarks>
    /// <param name="list">The list to apply.</param>
    /// <param name="level">
    /// The level to place the paragraph at, from <c>0</c> to <c>8</c>. Defaults to the first level.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="list"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="level"/> is not between 0 and 8.</exception>
    /// <exception cref="InvalidOfficeFileException">The list has no readable identifier.</exception>
    public void ApplyList(WordList list, int level = 0)
    {
        ArgumentNullException.ThrowIfNull(list);
        ListReference.EnsureValidLevel(level, nameof(level));

        // Read the identifier before touching the tree so that a malformed list leaves no partial markup.
        int listId = list.Id;
        ListLevel = level;
        ListId = listId;
    }

    /// <summary>
    /// Gets or sets the horizontal alignment of the paragraph.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the paragraph inherits its alignment from styles or defaults,
    /// which is exactly what Office Open XML expresses by omitting <c>w:jc</c>.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is not a defined alignment.</exception>
    /// <exception cref="InvalidOfficeFileException">A stored alignment cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The paragraph stores an alignment that Office Open XML defines but this version does not
    /// model (for example <c>mediumKashida</c>).
    /// </exception>
    public ParagraphAlignment? Alignment
    {
        get
        {
            EnumValue<W.JustificationValues>? stored = Properties?.GetFirstChild<W.Justification>()?.Val;
            return stored is null
                ? null
                : ToParagraphAlignment(ReadStoredValue(() => stored.Value, "alignment"));
        }

        set
        {
            if (value is null)
            {
                W.ParagraphProperties? properties = Properties;
                if (properties is not null)
                {
                    properties.RemoveAllChildren<W.Justification>();
                    PruneParagraphPropertiesIfEmpty();
                }

                return;
            }

            // Validate before touching the tree so that a rejected value leaves no partial markup.
            W.JustificationValues ooxml = ToOoxmlAlignment(value.Value);

            W.ParagraphProperties owned = EnsureParagraphProperties();
            owned.RemoveAllChildren<W.Justification>();
            InsertProperty(owned, new W.Justification { Val = ooxml });
        }
    }

    /// <summary>
    /// Gets or sets the first-line indent in points, or <see langword="null"/> when the paragraph
    /// does not set one.
    /// </summary>
    /// <remarks>
    /// Office Open XML stores indents in twentieths of a point (twips, <c>ST_SignedTwipsMeasure</c>),
    /// so an assigned value is rounded to the nearest <c>0.05</c> points. Attributes that this
    /// version does not model (for example <c>w:firstLineChars</c>) are preserved, and reading this
    /// property reports <see langword="null"/> when only those are present.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The assigned value is not a finite, non-negative number of points, or it is too large to be
    /// represented in twips (twentieths of a point).
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored indent cannot be parsed.</exception>
    public double? FirstLineIndentPoints
    {
        get => ReadTwips(Properties?.GetFirstChild<W.Indentation>()?.FirstLine, "first-line indent");
        set => SetIndent(value, nameof(value), "The first-line indent must be a finite, non-negative number of points.", "firstLine");
    }

    /// <summary>
    /// Gets or sets the left indent in points, or <see langword="null"/> when the paragraph does not
    /// set one.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The assigned value is not a finite, non-negative number of points, or it is too large to be
    /// represented in twips (twentieths of a point).
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored indent cannot be parsed.</exception>
    public double? LeftIndentPoints
    {
        get => ReadTwips(Properties?.GetFirstChild<W.Indentation>()?.Left, "left indent");
        set => SetIndent(value, nameof(value), "The left indent must be a finite, non-negative number of points.", "left");
    }

    /// <summary>
    /// Gets or sets the right indent in points, or <see langword="null"/> when the paragraph does not
    /// set one.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The assigned value is not a finite, non-negative number of points, or it is too large to be
    /// represented in twips (twentieths of a point).
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored indent cannot be parsed.</exception>
    public double? RightIndentPoints
    {
        get => ReadTwips(Properties?.GetFirstChild<W.Indentation>()?.Right, "right indent");
        set => SetIndent(value, nameof(value), "The right indent must be a finite, non-negative number of points.", "right");
    }

    /// <summary>
    /// Gets or sets the hanging indent in points, or <see langword="null"/> when the paragraph does
    /// not set one.
    /// </summary>
    /// <remarks>
    /// A hanging indent is combined with the left indent: Office Open XML stores it as
    /// <c>w:hanging</c>, a non-negative twips measure. As with the other indents, the assigned value
    /// is rounded to the nearest <c>0.05</c> points.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The assigned value is not a finite, non-negative number of points, or it is too large to be
    /// represented in twips (twentieths of a point).
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored indent cannot be parsed.</exception>
    public double? HangingIndentPoints
    {
        get => ReadTwips(Properties?.GetFirstChild<W.Indentation>()?.Hanging, "hanging indent");
        set => SetIndent(value, nameof(value), "The hanging indent must be a finite, non-negative number of points.", "hanging");
    }

    /// <summary>
    /// Gets or sets the spacing before the paragraph in points, or <see langword="null"/> when the
    /// paragraph does not set one.
    /// </summary>
    /// <remarks>
    /// Office Open XML stores spacing in twentieths of a point (twips), so an assigned value is
    /// rounded to the nearest <c>0.05</c> points. The <c>w:beforeLines</c> variant, which this
    /// version does not model, is preserved but is not reported here.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The assigned value is not a finite, non-negative number of points, or it is too large to be
    /// represented in twips (twentieths of a point).
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored spacing cannot be parsed.</exception>
    public double? SpacingBeforePoints
    {
        get => ReadTwips(Properties?.GetFirstChild<W.SpacingBetweenLines>()?.Before, "spacing before");
        set => SetSpacing(value, nameof(value), "The spacing before must be a finite, non-negative number of points.", Before: true);
    }

    /// <summary>
    /// Gets or sets the spacing after the paragraph in points, or <see langword="null"/> when the
    /// paragraph does not set one.
    /// </summary>
    /// <remarks>
    /// Office Open XML stores spacing in twentieths of a point (twips), so an assigned value is
    /// rounded to the nearest <c>0.05</c> points. The <c>w:afterLines</c> variant, which this version
    /// does not model, is preserved but is not reported here.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The assigned value is not a finite, non-negative number of points, or it is too large to be
    /// represented in twips (twentieths of a point).
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored spacing cannot be parsed.</exception>
    public double? SpacingAfterPoints
    {
        get => ReadTwips(Properties?.GetFirstChild<W.SpacingBetweenLines>()?.After, "spacing after");
        set => SetSpacing(value, nameof(value), "The spacing after must be a finite, non-negative number of points.", Before: false);
    }

    /// <summary>
    /// Gets or sets the exact line spacing of the paragraph in points, or <see langword="null"/>
    /// when the paragraph does not set one.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Office Open XML expresses line spacing in two different units, selected by <c>w:lineRule</c>:
    /// <c>exact</c> and <c>atLeast</c> use twentieths of a point, while <c>auto</c> (the default
    /// when the attribute is absent) uses a multiple of a line, where <c>240</c> means single
    /// spacing. Assigning this property always writes <c>w:lineRule="exact"</c>.
    /// </para>
    /// <para>
    /// Because a multiple of a line cannot be expressed in points, reading this property on a
    /// paragraph that stores <c>w:lineRule="auto"</c> throws
    /// <see cref="UnsupportedOfficeFeatureException"/> instead of returning a wrong number. The
    /// original markup is preserved.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The assigned value is not a finite, positive number of points, or it is too large to be
    /// represented in twips (twentieths of a point).
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored line spacing cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The paragraph stores line spacing as a multiple of a line (<c>w:lineRule="auto"</c>).
    /// </exception>
    public double? LineSpacingPoints
    {
        get
        {
            W.SpacingBetweenLines? spacing = Properties?.GetFirstChild<W.SpacingBetweenLines>();
            if (spacing?.Line is null)
            {
                return null;
            }

            EnumValue<W.LineSpacingRuleValues>? stored = spacing.LineRule;
            W.LineSpacingRuleValues rule = stored is null
                ? W.LineSpacingRuleValues.Auto
                : ReadStoredValue(() => stored.Value, "line spacing rule");

            if (rule == W.LineSpacingRuleValues.Auto)
            {
                throw new UnsupportedOfficeFeatureException(
                    "The paragraph stores line spacing as a multiple of a line (w:lineRule=\"auto\"), which this version does not model in points.");
            }

            return ReadTwips(spacing.Line, "line spacing");
        }

        set
        {
            if (value is null)
            {
                ClearSpacingAttribute("lineRule");
                ClearSpacingAttribute("line");
                PruneParagraphPropertiesIfEmpty();
                return;
            }

            double points = RequireFinitePositive(value, nameof(value), "The line spacing must be a finite, positive number of points.");
            W.ParagraphProperties properties = EnsureParagraphProperties();
            W.SpacingBetweenLines spacing = EnsureSpacing(properties);
            spacing.Line = PointsToTwips(points, nameof(value)).ToString(CultureInfo.InvariantCulture);
            spacing.LineRule = W.LineSpacingRuleValues.Exact;
            InsertProperty(properties, spacing);
        }
    }

    /// <summary>
    /// Gets the tab stops that this paragraph defines, in document order, or an empty list when it
    /// defines none of its own.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Office Open XML stores these in <c>w:pPr/w:tabs</c>. A paragraph also inherits tab stops from
    /// its style, from numbering and from the default tab interval; those sources are not modelled,
    /// so this list only reports the paragraph's own stops.
    /// </para>
    /// <para>
    /// The list is rebuilt on each access and reflects the current content of the paragraph.
    /// </para>
    /// </remarks>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// A stored stop uses an alignment that Office Open XML defines but this version does not model
    /// (for example <c>start</c>, <c>end</c> or <c>num</c>).
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">A stored stop is missing a required attribute or is malformed.</exception>
    public IReadOnlyList<TabStop> TabStops
    {
        get
        {
            W.Tabs? tabs = Properties?.GetFirstChild<W.Tabs>();
            if (tabs is null)
            {
                return Array.Empty<TabStop>();
            }

            List<TabStop> stops = new();
            foreach (W.TabStop stop in tabs.Elements<W.TabStop>())
            {
                stops.Add(ReadTabStop(stop));
            }

            return stops;
        }
    }

    /// <summary>
    /// Adds a tab stop to this paragraph.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The stop is appended to the paragraph's own list (<c>w:pPr/w:tabs</c>), which is created when
    /// it does not exist yet. Stops that this version does not model are left untouched.
    /// </para>
    /// <para>
    /// Stops are kept in insertion order; they are not re-sorted by position. The schema does not
    /// require an order, and Word orders them when it renders or edits the paragraph.
    /// </para>
    /// </remarks>
    /// <param name="tabStop">The tab stop to add.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The position of <paramref name="tabStop"/> is not finite, or it is too large to be
    /// represented in twips (twentieths of a point), or its alignment or leader is not a defined
    /// value.
    /// </exception>
    public void AddTabStop(TabStop tabStop)
    {
        // Validate before touching the tree so that a rejected value leaves no partial markup.
        int position = PointsToTwips(tabStop.PositionPoints, nameof(tabStop));
        W.TabStopValues alignment = ToOoxmlTabStopAlignment(tabStop.Alignment);
        W.TabStopLeaderCharValues? leader = ToOoxmlTabStopLeader(tabStop.Leader);

        W.ParagraphProperties properties = EnsureParagraphProperties();
        W.TabStop created = new()
        {
            Val = alignment,
            Position = position,
        };

        if (leader is not null)
        {
            created.Leader = leader;
        }

        W.Tabs tabs = EnsureTabs(properties);
        tabs.AppendChild(created);
        InsertProperty(properties, tabs);
    }

    /// <summary>
    /// Removes every tab stop that this paragraph defines.
    /// </summary>
    /// <remarks>
    /// Office Open XML does not allow an empty <c>w:tabs</c> element, so it is removed as a whole.
    /// Inherited tab stops (style, numbering and the default interval) are not affected.
    /// </remarks>
    public void ClearTabStops()
    {
        W.ParagraphProperties? properties = Properties;
        if (properties is null)
        {
            return;
        }

        properties.RemoveAllChildren<W.Tabs>();
        PruneParagraphPropertiesIfEmpty();
    }

    /// <inheritdoc />
    public override string ToString() => Text;

    private static T ReadStoredValue<T>(Func<T> read, string propertyDescription)
    {
        try
        {
            return read();
        }
        catch (FormatException exception)
        {
            throw new InvalidOfficeFileException(
                $"A paragraph stores a {propertyDescription} value that this document cannot read.",
                exception);
        }
    }

    private static double? ReadTwips(StringValue? stored, string propertyDescription)
    {
        if (stored is null)
        {
            return null;
        }

        if (!int.TryParse(stored.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips))
        {
            throw new InvalidOfficeFileException($"A paragraph stores a {propertyDescription} of '{stored.Value}' that this document cannot read.");
        }

        return TwipsToPoints(twips);
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
            $"The paragraph stores the alignment '{value}', a value that Office Open XML defines but this version does not model.");
    }

    private static W.JustificationValues ToOoxmlAlignment(ParagraphAlignment value)
    {
        if (value == ParagraphAlignment.Left)
        {
            return W.JustificationValues.Left;
        }

        if (value == ParagraphAlignment.Center)
        {
            return W.JustificationValues.Center;
        }

        if (value == ParagraphAlignment.Right)
        {
            return W.JustificationValues.Right;
        }

        if (value == ParagraphAlignment.Justify)
        {
            return W.JustificationValues.Distribute;
        }

        if (value == ParagraphAlignment.Both)
        {
            return W.JustificationValues.Both;
        }

        throw new ArgumentOutOfRangeException(nameof(value), value, "The alignment is not a defined value.");
    }

    private void SetIndent(double? value, string parameterName, string message, string attributeLocalName)
    {
        if (value is null)
        {
            ClearIndentAttribute(attributeLocalName);
            PruneParagraphPropertiesIfEmpty();
            return;
        }

        double points = RequireFiniteNonNegative(value, parameterName, message);
        W.ParagraphProperties properties = EnsureParagraphProperties();
        W.Indentation indentation = EnsureIndentation(properties);
        string twips = PointsToTwips(points, parameterName).ToString(CultureInfo.InvariantCulture);
        switch (attributeLocalName)
        {
            case "firstLine":
                indentation.FirstLine = twips;
                break;
            case "left":
                indentation.Left = twips;
                break;
            case "right":
                indentation.Right = twips;
                break;
            default:
                indentation.Hanging = twips;
                break;
        }

        InsertProperty(properties, indentation);
    }

    private void SetSpacing(double? value, string parameterName, string message, bool Before)
    {
        if (value is null)
        {
            ClearSpacingAttribute(Before ? "before" : "after");
            PruneParagraphPropertiesIfEmpty();
            return;
        }

        double points = RequireFiniteNonNegative(value, parameterName, message);
        W.ParagraphProperties properties = EnsureParagraphProperties();
        W.SpacingBetweenLines spacing = EnsureSpacing(properties);
        string twips = PointsToTwips(points, parameterName).ToString(CultureInfo.InvariantCulture);
        if (Before)
        {
            spacing.Before = twips;
        }
        else
        {
            spacing.After = twips;
        }

        InsertProperty(properties, spacing);
    }

    private static double RequireFiniteNonNegative(double? points, string parameterName, string message)
    {
        if (points is not { } assigned || !double.IsFinite(assigned) || assigned < 0.0)
        {
            throw new ArgumentOutOfRangeException(parameterName, points, message);
        }

        return assigned;
    }

    private static double RequireFinitePositive(double? points, string parameterName, string message)
    {
        if (points is not { } assigned || !double.IsFinite(assigned) || assigned <= 0.0)
        {
            throw new ArgumentOutOfRangeException(parameterName, points, message);
        }

        return assigned;
    }

    private static int PointsToTwips(double points, string parameterName)
    {
        double twips = Math.Round(points * 20.0, MidpointRounding.AwayFromZero);
        if (twips < int.MinValue || twips > int.MaxValue)
        {
            // Office Open XML stores these measures as a 32-bit twips value. An unchecked cast would
            // silently turn a huge number of points into a nonsense (often negative) measure.
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The value is too large to be represented in twips (twentieths of a point).");
        }

        return (int)twips;
    }

    private static double TwipsToPoints(int twips)
        => twips / 20.0;

    private W.ParagraphProperties EnsureParagraphProperties()
    {
        W.ParagraphProperties? properties = _element.GetFirstChild<W.ParagraphProperties>();
        if (properties is null)
        {
            properties = new W.ParagraphProperties();
            _element.InsertAt(properties, 0);
        }

        return properties;
    }

    private void PruneParagraphPropertiesIfEmpty()
    {
        W.ParagraphProperties? properties = _element.GetFirstChild<W.ParagraphProperties>();
        if (properties is not null && !properties.HasChildren && !properties.GetAttributes().Any())
        {
            properties.Remove();
        }
    }

    private void RemoveNumberingChild<T>()
        where T : OpenXmlElement
    {
        W.NumberingProperties? numbering = Properties?.GetFirstChild<W.NumberingProperties>();
        if (numbering is null)
        {
            return;
        }

        numbering.RemoveAllChildren<T>();
        if (!numbering.HasChildren && !numbering.GetAttributes().Any())
        {
            numbering.Remove();
        }

        PruneParagraphPropertiesIfEmpty();
    }

    private void ClearIndentAttribute(string attributeLocalName)
    {
        W.Indentation? indentation = _element.GetFirstChild<W.ParagraphProperties>()?.GetFirstChild<W.Indentation>();
        if (indentation is null)
        {
            return;
        }

        indentation.RemoveAttribute(attributeLocalName, WordprocessingNamespace);
        if (!indentation.GetAttributes().Any())
        {
            indentation.Remove();
        }
    }

    private void ClearSpacingAttribute(string attributeLocalName)
    {
        W.SpacingBetweenLines? spacing = _element.GetFirstChild<W.ParagraphProperties>()?.GetFirstChild<W.SpacingBetweenLines>();
        if (spacing is null)
        {
            return;
        }

        spacing.RemoveAttribute(attributeLocalName, WordprocessingNamespace);
        if (!spacing.GetAttributes().Any())
        {
            spacing.Remove();
        }
    }

    private static W.Indentation EnsureIndentation(W.ParagraphProperties properties)
    {
        W.Indentation? indentation = properties.GetFirstChild<W.Indentation>();
        if (indentation is null)
        {
            indentation = new W.Indentation();
            properties.AppendChild(indentation);
        }

        return indentation;
    }

    private static W.SpacingBetweenLines EnsureSpacing(W.ParagraphProperties properties)
    {
        W.SpacingBetweenLines? spacing = properties.GetFirstChild<W.SpacingBetweenLines>();
        if (spacing is null)
        {
            spacing = new W.SpacingBetweenLines();
            properties.AppendChild(spacing);
        }

        return spacing;
    }

    private static W.Tabs EnsureTabs(W.ParagraphProperties properties)
    {
        W.Tabs? tabs = properties.GetFirstChild<W.Tabs>();
        if (tabs is null)
        {
            tabs = new W.Tabs();
            properties.AppendChild(tabs);
        }

        return tabs;
    }

    private static W.NumberingProperties EnsureNumberingProperties(W.ParagraphProperties properties)
    {
        W.NumberingProperties? numbering = properties.GetFirstChild<W.NumberingProperties>();
        if (numbering is null)
        {
            numbering = new W.NumberingProperties();
            InsertProperty(properties, numbering);
        }

        return numbering;
    }

    private static TabStop ReadTabStop(W.TabStop stored)
    {
        EnumValue<W.TabStopValues>? alignment = stored.Val;
        if (alignment is null)
        {
            throw new InvalidOfficeFileException("A paragraph stores a tab stop without the required w:val attribute.");
        }

        Int32Value? position = stored.Position;
        if (position is null)
        {
            throw new InvalidOfficeFileException("A paragraph stores a tab stop without the required w:pos attribute.");
        }

        TabStopAlignment mappedAlignment = ToTabStopAlignment(ReadStoredValue(() => alignment.Value, "tab stop alignment"));
        int twips = ReadStoredValue(() => position.Value, "tab stop position");

        TabStopLeader leader = TabStopLeader.None;
        EnumValue<W.TabStopLeaderCharValues>? storedLeader = stored.Leader;
        if (storedLeader is not null)
        {
            leader = ToTabStopLeader(ReadStoredValue(() => storedLeader.Value, "tab stop leader"));
        }

        return new TabStop(TwipsToPoints(twips), mappedAlignment, leader);
    }

    private static TabStopAlignment ToTabStopAlignment(W.TabStopValues value)
        => TabStopsFromOoxml.TryGetValue(value, out TabStopAlignment alignment)
            ? alignment
            : throw new UnsupportedOfficeFeatureException(
                $"The paragraph stores the tab stop alignment '{value}', a value that Office Open XML defines but this version does not model.");

    private static TabStopLeader ToTabStopLeader(W.TabStopLeaderCharValues value)
        => TabStopLeadersFromOoxml.TryGetValue(value, out TabStopLeader leader)
            ? leader
            : throw new InvalidOfficeFileException(
                $"The paragraph stores the tab stop leader '{value}' that this document cannot read.");

    private static readonly Dictionary<W.TabStopValues, TabStopAlignment> TabStopsFromOoxml =
        new Dictionary<W.TabStopValues, TabStopAlignment>
        {
            [W.TabStopValues.Clear] = TabStopAlignment.Clear,
            [W.TabStopValues.Left] = TabStopAlignment.Left,
            [W.TabStopValues.Center] = TabStopAlignment.Center,
            [W.TabStopValues.Right] = TabStopAlignment.Right,
            [W.TabStopValues.Decimal] = TabStopAlignment.Decimal,
            [W.TabStopValues.Bar] = TabStopAlignment.Bar,
        };

    private static readonly Dictionary<W.TabStopLeaderCharValues, TabStopLeader> TabStopLeadersFromOoxml =
        new Dictionary<W.TabStopLeaderCharValues, TabStopLeader>
        {
            [W.TabStopLeaderCharValues.None] = TabStopLeader.None,
            [W.TabStopLeaderCharValues.Dot] = TabStopLeader.Dot,
            [W.TabStopLeaderCharValues.Hyphen] = TabStopLeader.Hyphen,
            [W.TabStopLeaderCharValues.Underscore] = TabStopLeader.Underscore,
            [W.TabStopLeaderCharValues.Heavy] = TabStopLeader.Heavy,
            [W.TabStopLeaderCharValues.MiddleDot] = TabStopLeader.MiddleDot,
        };

    private static W.TabStopValues ToOoxmlTabStopAlignment(TabStopAlignment value)
    {
        if (value == TabStopAlignment.Clear)
        {
            return W.TabStopValues.Clear;
        }

        if (value == TabStopAlignment.Left)
        {
            return W.TabStopValues.Left;
        }

        if (value == TabStopAlignment.Center)
        {
            return W.TabStopValues.Center;
        }

        if (value == TabStopAlignment.Right)
        {
            return W.TabStopValues.Right;
        }

        if (value == TabStopAlignment.Decimal)
        {
            return W.TabStopValues.Decimal;
        }

        if (value == TabStopAlignment.Bar)
        {
            return W.TabStopValues.Bar;
        }

        throw new ArgumentOutOfRangeException(nameof(value), value, "The tab stop alignment is not a defined value.");
    }

    private static W.TabStopLeaderCharValues? ToOoxmlTabStopLeader(TabStopLeader value)
    {
        if (value == TabStopLeader.None)
        {
            // @w:leader is omitted rather than written as "none", which is what Word does.
            return null;
        }

        if (value == TabStopLeader.Dot)
        {
            return W.TabStopLeaderCharValues.Dot;
        }

        if (value == TabStopLeader.Hyphen)
        {
            return W.TabStopLeaderCharValues.Hyphen;
        }

        if (value == TabStopLeader.Underscore)
        {
            return W.TabStopLeaderCharValues.Underscore;
        }

        if (value == TabStopLeader.Heavy)
        {
            return W.TabStopLeaderCharValues.Heavy;
        }

        if (value == TabStopLeader.MiddleDot)
        {
            return W.TabStopLeaderCharValues.MiddleDot;
        }

        throw new ArgumentOutOfRangeException(nameof(value), value, "The tab stop leader is not a defined value.");
    }

    /// <summary>
    /// Inserts a property element at the position required by <c>CT_PPr</c>, detaching it first when
    /// it is already in the tree. This is used instead of <c>AppendChild</c> so a paragraph that
    /// already carries markup this version does not model keeps its <c>w:pPr</c> children in schema
    /// order.
    /// </summary>
    private static void InsertProperty(W.ParagraphProperties properties, OpenXmlElement element)
        => InsertInSchemaOrder(properties, element, ParagraphPropertyOrder);

    /// <summary>
    /// Inserts a child of <c>w:numPr</c> at the position required by <c>CT_NumPr</c>, preserving any
    /// numbering change markup (<c>w:numberingChange</c>/<c>w:ins</c>) this version does not model.
    /// </summary>
    private static void InsertNumberingProperty(W.NumberingProperties properties, OpenXmlElement element)
        => InsertInSchemaOrder(properties, element, NumberingPropertyOrder);

    private static void InsertInSchemaOrder(OpenXmlCompositeElement parent, OpenXmlElement element, string[] order)
    {
        if (element.Parent is not null)
        {
            element.Remove();
        }

        int rank = RankOf(order, element.LocalName);
        foreach (OpenXmlElement child in parent.ChildElements)
        {
            if (RankOf(order, child.LocalName) > rank)
            {
                parent.InsertBefore(element, child);
                return;
            }
        }

        parent.AppendChild(element);
    }

    private static int RankOf(string[] order, string localName)
    {
        int index = Array.IndexOf(order, localName);
        return index < 0 ? order.Length : index;
    }

    /// <summary>
    /// Removes every child except the paragraph properties (<c>w:pPr</c>), which are kept so that
    /// replacing text does not silently discard alignment, indentation, spacing or styles.
    /// </summary>
    private void RemoveContentKeepingParagraphProperties()
    {
        foreach (OpenXmlElement child in _element.ChildElements.ToArray())
        {
            if (child is not W.ParagraphProperties)
            {
                child.Remove();
            }
        }
    }

    private static W.Run CreateRun(string text)
        => new(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });
}
