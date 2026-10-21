using OfficeEngine.OpenXml;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PageSizeValue = OfficeEngine.Word.PageSize;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over a section of a <see cref="WordDocument"/> (<c>w:sectPr</c>).
/// </summary>
/// <remarks>
/// <para>
/// In WordprocessingML every section ends with a set of section properties. The properties of the
/// last section are the <c>w:sectPr</c> that closes the body; every earlier section stores its
/// properties in the paragraph that ends it (<c>w:pPr/w:sectPr</c>). <see cref="WordDocument.Sections"/>
/// returns them in document order and <see cref="WordDocument.AddSection"/> appends a new one.
/// </para>
/// <para>
/// This version models the page size and orientation, the page margins, how the section starts and its
/// headers and footers. Columns, borders, line numbering, page numbering, document grid and the rest of
/// the section properties are preserved but not exposed.
/// </para>
/// <para>
/// This is a lightweight wrapper: it does not own the underlying content and must not be used after
/// the <see cref="WordDocument"/> that produced it has been disposed.
/// </para>
/// </remarks>
public sealed class Section
{
    // Child order required by CT_SectPr. The names are the XML local names and the sequence is the
    // one the Open XML SDK generates for a SectionProperties element. Properties are inserted
    // following this table instead of being appended, so a section that already carries markup this
    // version does not model (columns, headers, a document grid) stays valid.
    private static readonly string[] SectionPropertyOrder =
    [
        "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
        "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign",
        "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
        "sectPrChange",
    ];

    private readonly W.SectionProperties _element;
    private readonly MainDocumentPart _mainPart;

    internal Section(W.SectionProperties element, MainDocumentPart mainPart)
    {
        _element = element;
        _mainPart = mainPart;
    }

    /// <summary>
    /// Gets or sets the page size and orientation of the section.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the section does not state a size and the application applies its
    /// own default, which is exactly what Office Open XML expresses by omitting <c>w:pgSz</c>.
    /// Assigning <see langword="null"/> removes the element. The orientation is stored separately and
    /// does not swap the dimensions; use <see cref="OfficeEngine.Word.PageSize.Rotate"/> for that.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A dimension or the orientation of the assigned value is not valid.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">
    /// The stored <c>w:pgSz</c> lacks a required dimension or stores one that cannot be parsed.
    /// </exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The stored orientation is a value Office Open XML defines but this version does not model.
    /// </exception>
    public PageSizeValue? PageSize
    {
        get => PageSizeValue.Read(_element.GetFirstChild<W.PageSize>(), "section");
        set
        {
            if (value is null)
            {
                _element.RemoveAllChildren<W.PageSize>();
                return;
            }

            W.PageSize element = _element.GetFirstChild<W.PageSize>() ?? new W.PageSize();
            value.Value.ApplyTo(element);
            SchemaOrder.Insert(_element, element, SectionPropertyOrder);
        }
    }

    /// <summary>
    /// Gets or sets the page margins of the section, in points.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the section does not state margins, which is what Office Open XML
    /// expresses by omitting <c>w:pgMar</c>. Assigning <see langword="null"/> removes the element.
    /// The header and footer distances and the gutter are not modelled: when the element already
    /// exists, they are preserved as they are.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">A margin of the assigned value is not valid.</exception>
    /// <exception cref="InvalidOfficeFileException">
    /// The stored <c>w:pgMar</c> lacks one of the four required margins or stores one that cannot be
    /// parsed.
    /// </exception>
    public PageMargins? Margins
    {
        get => PageMargins.Read(_element.GetFirstChild<W.PageMargin>(), "section");
        set
        {
            if (value is null)
            {
                _element.RemoveAllChildren<W.PageMargin>();
                return;
            }

            W.PageMargin element = _element.GetFirstChild<W.PageMargin>() ?? new W.PageMargin();
            value.Value.ApplyTo(element);
            SchemaOrder.Insert(_element, element, SectionPropertyOrder);
        }
    }

    /// <summary>
    /// Gets or sets how the section starts.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the section does not state it, which is exactly what Office Open
    /// XML expresses by omitting <c>w:type</c>; the default is <see cref="SectionStartType.NextPage"/>.
    /// Assigning <see langword="null"/> removes the element.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is not a defined start type.</exception>
    /// <exception cref="InvalidOfficeFileException">
    /// The stored <c>w:type</c> lacks its value or stores one that cannot be parsed.
    /// </exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The stored value is one Office Open XML defines but this version does not model.
    /// </exception>
    public SectionStartType? StartType
    {
        get
        {
            W.SectionType? stored = _element.GetFirstChild<W.SectionType>();
            if (stored is null)
            {
                return null;
            }

            EnumValue<W.SectionMarkValues>? value = stored.Val
                ?? throw new InvalidOfficeFileException("A section stores a w:type element without the required w:val attribute.");

            W.SectionMarkValues mark;
            try
            {
                mark = value.Value;
            }
            catch (FormatException exception)
            {
                throw new InvalidOfficeFileException(
                    "A section stores a w:type value that this document cannot read.",
                    exception);
            }

            if (mark == W.SectionMarkValues.NextPage)
            {
                return SectionStartType.NextPage;
            }

            if (mark == W.SectionMarkValues.Continuous)
            {
                return SectionStartType.Continuous;
            }

            if (mark == W.SectionMarkValues.EvenPage)
            {
                return SectionStartType.EvenPage;
            }

            if (mark == W.SectionMarkValues.OddPage)
            {
                return SectionStartType.OddPage;
            }

            if (mark == W.SectionMarkValues.NextColumn)
            {
                return SectionStartType.NextColumn;
            }

            throw new UnsupportedOfficeFeatureException(
                $"A section stores the start type '{mark}', a value that Office Open XML defines but this version does not model.");
        }

        set
        {
            if (value is null)
            {
                _element.RemoveAllChildren<W.SectionType>();
                return;
            }

            // Validate before touching the tree so that a rejected value leaves no partial markup.
            W.SectionMarkValues mark = ToOoxmlStartType(value.Value);
            _element.RemoveAllChildren<W.SectionType>();
            WriteStartType(_element, mark);
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether the first page of the section uses a different
    /// header and footer (<c>w:titlePg</c>).
    /// </summary>
    /// <remarks>
    /// Word displays the <see cref="HeaderFooterType.FirstPage"/> header or footer only when this is
    /// <see langword="true"/>. The property has no effect on the other types. Assigning
    /// <see langword="false"/> removes the element, which is exactly what Office Open XML expresses by
    /// omitting it.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">
    /// The stored <c>w:titlePg</c> stores a value that cannot be parsed as a boolean.
    /// </exception>
    public bool DifferentFirstPage
    {
        get
        {
            W.TitlePage? stored = _element.GetFirstChild<W.TitlePage>();
            if (stored is null)
            {
                return false;
            }

            if (stored.Val is null)
            {
                return true;
            }

            try
            {
                return stored.Val.Value;
            }
            catch (FormatException exception)
            {
                throw new InvalidOfficeFileException(
                    "A section stores a w:titlePg value that this document cannot read.",
                    exception);
            }
        }

        set
        {
            _element.RemoveAllChildren<W.TitlePage>();
            if (value)
            {
                SchemaOrder.Insert(_element, new W.TitlePage(), SectionPropertyOrder);
            }
        }
    }

    /// <summary>
    /// Gets the header of the given type, or <see langword="null"/> when the section does not state one.
    /// </summary>
    /// <remarks>
    /// <see langword="null"/> means the section inherits the header of the previous section (Word's
    /// "link to previous"), which is exactly what Office Open XML expresses by omitting the
    /// reference. Use <see cref="AddHeader"/> to give the section its own header.
    /// </remarks>
    /// <param name="type">Which pages the header applies to.</param>
    /// <returns>The header, or <see langword="null"/> when there is none.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="type"/> is not a defined value.</exception>
    /// <exception cref="InvalidOfficeFileException">
    /// A stored header reference lacks its type or its relationship, or the part it names is missing or
    /// is not a header.
    /// </exception>
    public HeaderFooter? Header(HeaderFooterType type) => ReadHeaderFooter(type, header: true);

    /// <summary>
    /// Gets the footer of the given type, or <see langword="null"/> when the section does not state one.
    /// </summary>
    /// <remarks>See <see cref="Header"/> for the meaning of <see langword="null"/> and the failures.</remarks>
    /// <param name="type">Which pages the footer applies to.</param>
    /// <returns>The footer, or <see langword="null"/> when there is none.</returns>
    public HeaderFooter? Footer(HeaderFooterType type) => ReadHeaderFooter(type, header: false);

    /// <summary>
    /// Gets the header of the given type, creating it when the section does not have one.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A new header is an empty paragraph in a new part that the section references, so the section no
    /// longer inherits the header of the previous section. Calling this for a type the section already
    /// has returns the existing header; it never replaces it.
    /// </para>
    /// <para>
    /// A referenced part is shared when several sections point to it, so editing it from one section
    /// changes what the others show. That is what Word's "link to previous" writes.
    /// </para>
    /// </remarks>
    /// <param name="type">Which pages the header applies to.</param>
    /// <returns>The existing or new header.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="type"/> is not a defined value.</exception>
    /// <exception cref="InvalidOfficeFileException">The document has no main document part.</exception>
    public HeaderFooter AddHeader(HeaderFooterType type) => AddHeaderFooter(type, header: true);

    /// <summary>
    /// Gets the footer of the given type, creating it when the section does not have one.
    /// </summary>
    /// <remarks>See <see cref="AddHeader"/> for what creating one does.</remarks>
    /// <param name="type">Which pages the footer applies to.</param>
    /// <returns>The existing or new footer.</returns>
    public HeaderFooter AddFooter(HeaderFooterType type) => AddHeaderFooter(type, header: false);

    /// <summary>
    /// Removes the header of the given type from the section.
    /// </summary>
    /// <remarks>
    /// The part is not deleted: another section may still reference it, and Office Open XML does not
    /// require every part to be referenced. The section goes back to inheriting the header of the
    /// previous section. Removing a type the section does not have does nothing.
    /// </remarks>
    /// <param name="type">Which pages the header applies to.</param>
    /// <returns><see langword="true"/> when a reference was removed; otherwise, <see langword="false"/>.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="type"/> is not a defined value.</exception>
    /// <exception cref="InvalidOfficeFileException">A stored reference has a type that cannot be read.</exception>
    public bool RemoveHeader(HeaderFooterType type) => RemoveHeaderFooter(type, header: true);

    /// <summary>
    /// Removes the footer of the given type from the section.
    /// </summary>
    /// <remarks>See <see cref="RemoveHeader"/> for what removing one does.</remarks>
    /// <param name="type">Which pages the footer applies to.</param>
    /// <returns><see langword="true"/> when a reference was removed; otherwise, <see langword="false"/>.</returns>
    public bool RemoveFooter(HeaderFooterType type) => RemoveHeaderFooter(type, header: false);

    private HeaderFooter? ReadHeaderFooter(HeaderFooterType type, bool header)
    {
        W.HeaderFooterValues wanted = ToOoxmlHeaderFooterType(type);
        IEnumerable<OpenXmlElement> references = header
            ? _element.Elements<W.HeaderReference>()
            : _element.Elements<W.FooterReference>();

        foreach (OpenXmlElement reference in references)
        {
            if (ReadReferenceType(reference, header) != wanted)
            {
                continue;
            }

            string id = ReadReferenceId(reference, header);
            if (!_mainPart.TryGetPartById(id, out OpenXmlPart? part) || part is null)
            {
                throw new InvalidOfficeFileException(
                    $"A section references a header or footer part ('{id}') that the document does not contain.");
            }

            if (header)
            {
                if (part is not HeaderPart headerPart)
                {
                    throw new InvalidOfficeFileException(
                        $"A section references '{id}' as a header, but it is not a readable header part.");
                }

                W.Header headerRoot = OoxmlPartReader.Read("the header part", () => headerPart.Header);
                return new HeaderFooter(headerRoot, isHeader: true);
            }

            if (part is not FooterPart footerPart)
            {
                throw new InvalidOfficeFileException(
                    $"A section references '{id}' as a footer, but it is not a readable footer part.");
            }

            W.Footer footerRoot = OoxmlPartReader.Read("the footer part", () => footerPart.Footer);
            return new HeaderFooter(footerRoot, isHeader: false);
        }

        return null;
    }

    private HeaderFooter AddHeaderFooter(HeaderFooterType type, bool header)
    {
        HeaderFooter? existing = ReadHeaderFooter(type, header);
        if (existing is not null)
        {
            return existing;
        }

        W.HeaderFooterValues ooxml = ToOoxmlHeaderFooterType(type);
        if (header)
        {
            HeaderPart part = _mainPart.AddNewPart<HeaderPart>();
            part.Header = new W.Header(new W.Paragraph());
            SchemaOrder.Insert(
                _element,
                new W.HeaderReference { Type = ooxml, Id = _mainPart.GetIdOfPart(part) },
                SectionPropertyOrder);
            return new HeaderFooter(part.Header, isHeader: true);
        }

        FooterPart footerPart = _mainPart.AddNewPart<FooterPart>();
        footerPart.Footer = new W.Footer(new W.Paragraph());
        SchemaOrder.Insert(
            _element,
            new W.FooterReference { Type = ooxml, Id = _mainPart.GetIdOfPart(footerPart) },
            SectionPropertyOrder);
        return new HeaderFooter(footerPart.Footer, isHeader: false);
    }

    private bool RemoveHeaderFooter(HeaderFooterType type, bool header)
    {
        W.HeaderFooterValues wanted = ToOoxmlHeaderFooterType(type);
        IEnumerable<OpenXmlElement> references = header
            ? _element.Elements<W.HeaderReference>()
            : _element.Elements<W.FooterReference>();

        List<OpenXmlElement> matches = [];
        foreach (OpenXmlElement reference in references)
        {
            if (ReadReferenceType(reference, header) == wanted)
            {
                matches.Add(reference);
            }
        }

        foreach (OpenXmlElement reference in matches)
        {
            reference.Remove();
        }

        return matches.Count > 0;
    }

    private static W.HeaderFooterValues ReadReferenceType(OpenXmlElement reference, bool header)
    {
        EnumValue<W.HeaderFooterValues>? type = header
            ? ((W.HeaderReference)reference).Type
            : ((W.FooterReference)reference).Type;

        if (type is null)
        {
            throw new InvalidOfficeFileException(
                "A section stores a header or footer reference without the required w:type attribute.");
        }

        try
        {
            return type.Value;
        }
        catch (FormatException exception)
        {
            throw new InvalidOfficeFileException(
                "A section stores a header or footer reference type that this document cannot read.",
                exception);
        }
    }

    private static string ReadReferenceId(OpenXmlElement reference, bool header)
    {
        StringValue? id = header
            ? ((W.HeaderReference)reference).Id
            : ((W.FooterReference)reference).Id;

        if (id is null || string.IsNullOrEmpty(id.Value))
        {
            throw new InvalidOfficeFileException(
                "A section stores a header or footer reference without the relationship it names.");
        }

        return id.Value!;
    }

    /// <summary>
    /// Creates the section properties this version writes for a document that does not state any:
    /// US Letter in portrait orientation and one-inch margins.
    /// </summary>
    internal static W.SectionProperties CreateDefault()
        => new(
            new W.PageSize { Width = PageSizeValue.Letter.WidthTwips, Height = PageSizeValue.Letter.HeightTwips },
            new W.PageMargin
            {
                Top = 1440,
                Right = 1440U,
                Bottom = 1440,
                Left = 1440U,
                Header = 720U,
                Footer = 720U,
                Gutter = 0U,
            });

    internal static void WriteStartType(W.SectionProperties element, W.SectionMarkValues mark)
        => SchemaOrder.Insert(element, new W.SectionType { Val = mark }, SectionPropertyOrder);

    private static W.HeaderFooterValues ToOoxmlHeaderFooterType(HeaderFooterType value)
    {
        if (value == HeaderFooterType.Primary)
        {
            return W.HeaderFooterValues.Default;
        }

        if (value == HeaderFooterType.FirstPage)
        {
            return W.HeaderFooterValues.First;
        }

        if (value == HeaderFooterType.EvenPage)
        {
            return W.HeaderFooterValues.Even;
        }

        throw new ArgumentOutOfRangeException(nameof(value), value, "The header or footer type is not a defined value.");
    }

    internal static W.SectionMarkValues ToOoxmlStartType(SectionStartType value)
    {
        if (value == SectionStartType.NextPage)
        {
            return W.SectionMarkValues.NextPage;
        }

        if (value == SectionStartType.Continuous)
        {
            return W.SectionMarkValues.Continuous;
        }

        if (value == SectionStartType.EvenPage)
        {
            return W.SectionMarkValues.EvenPage;
        }

        if (value == SectionStartType.OddPage)
        {
            return W.SectionMarkValues.OddPage;
        }

        if (value == SectionStartType.NextColumn)
        {
            return W.SectionMarkValues.NextColumn;
        }

        throw new ArgumentOutOfRangeException(nameof(value), value, "The section start type is not a defined value.");
    }
}
