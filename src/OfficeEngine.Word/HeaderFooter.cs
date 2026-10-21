using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over the header or footer of a section (<c>w:hdr</c> or <c>w:ftr</c>).
/// </summary>
/// <remarks>
/// <para>
/// A header or a footer is a part of the package that holds block-level content, the same content the
/// body holds: paragraphs and tables. A section points to it through a <c>w:headerReference</c> or a
/// <c>w:footerReference</c> in <c>w:sectPr</c>, and several sections can point to the same part: Word
/// calls that sharing "link to previous". Editing this view changes the part, so every section that
/// points to it sees the change; that is exactly what the format expresses.
/// </para>
/// <para>
/// This version models the paragraphs and tables of the header or footer and their text. Content it
/// does not model (images, content controls) is preserved but cannot be edited through the API.
/// </para>
/// <para>
/// This is a lightweight wrapper: it does not own the underlying content and must not be used after
/// the <see cref="WordDocument"/> that produced it has been disposed.
/// </para>
/// </remarks>
public sealed class HeaderFooter
{
    private readonly OpenXmlCompositeElement _root;

    internal HeaderFooter(OpenXmlCompositeElement root, bool isHeader)
    {
        _root = root;
        IsHeader = isHeader;
    }

    /// <summary>Gets a value indicating whether this is a header; otherwise it is a footer.</summary>
    public bool IsHeader { get; }

    /// <summary>Gets a live snapshot of the top-level paragraphs of the header or footer, in order.</summary>
    /// <remarks>
    /// A header or footer usually holds a single paragraph, but the schema allows any number of
    /// paragraphs and tables. Only direct children are reported, matching what the API can edit.
    /// </remarks>
    public IReadOnlyList<Paragraph> Paragraphs
        => _root.Elements<W.Paragraph>().Select(element => new Paragraph(element)).ToArray();

    /// <summary>Gets a live snapshot of the top-level tables of the header or footer, in order.</summary>
    /// <remarks>Only direct children are reported, matching what the API can edit.</remarks>
    public IReadOnlyList<Table> Tables
        => _root.Elements<W.Table>().Select(element => new Table(element)).ToArray();

    /// <summary>
    /// Gets the visible text of the header or footer: one line per top-level paragraph and one line per
    /// table row, joined by line-feed characters.
    /// </summary>
    /// <remarks>
    /// See <see cref="WordDocument.Text"/> for the shape of the text and for the content this version
    /// does not read. It is preserved when the document is saved.
    /// </remarks>
    public string Text => BlockText.Of(_root.ChildElements);

    /// <summary>Adds a new paragraph with the given text at the end of the header or footer.</summary>
    /// <param name="text">The text of the paragraph. May be empty.</param>
    /// <returns>A live view over the new paragraph.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="text"/> contains a tab, a line break or a character that Office Open XML cannot
    /// store. Add tabs and breaks with <see cref="Run.AddTab"/> and <see cref="Run.AddBreak"/>.
    /// </exception>
    public Paragraph AddParagraph(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        PlainText.EnsureValid(text, nameof(text));

        W.Paragraph element = new();
        if (text.Length > 0)
        {
            element.AppendChild(new W.Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        _root.AppendChild(element);
        return new Paragraph(element);
    }

    /// <summary>Adds a new paragraph with a page number field at the end of the header or footer.</summary>
    /// <param name="cachedResult">
    /// The result stored with the field, for example <c>1</c>. It is what a reader shows until Word
    /// recalculates the field.
    /// </param>
    /// <returns>A live view over the new paragraph.</returns>
    /// <remarks>
    /// The field is written, not calculated: Word recomputes the page number when the document is
    /// opened or printed. Use <see cref="Paragraph.AddField"/> for other fields, such as
    /// <c>NUMPAGES</c>.
    /// </remarks>
    public Paragraph AddPageNumber(string? cachedResult = null)
    {
        IReadOnlyList<Paragraph> paragraphs = Paragraphs;
        Paragraph? paragraph = null;
        if (paragraphs.Count > 0 && paragraphs[^1].Text.Length == 0 && paragraphs[^1].Runs.Count == 0)
        {
            // The part is created with an empty paragraph; the page number goes there instead of
            // leaving a blank line above it.
            paragraph = paragraphs[^1];
        }

        paragraph ??= AddParagraph(string.Empty);
        paragraph.AddField("PAGE", cachedResult);
        return paragraph;
    }

    /// <summary>Adds a new table with the given number of rows and columns at the end of the header or footer.</summary>
    /// <remarks>
    /// Every cell starts with a single empty paragraph. See <see cref="WordDocument.AddTable"/> for the
    /// table this version creates.
    /// </remarks>
    /// <param name="rows">The number of rows. Must be at least one.</param>
    /// <param name="columns">The number of columns. Must be at least one.</param>
    /// <returns>A live view over the new table.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="rows"/> or <paramref name="columns"/> is less than one.
    /// </exception>
    public Table AddTable(int rows, int columns)
    {
        if (rows < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(rows), rows, "A table must have at least one row.");
        }

        if (columns < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(columns), columns, "A table must have at least one column.");
        }

        W.Table element = Table.Create(rows, columns);
        _root.AppendChild(element);
        return new Table(element);
    }
}
