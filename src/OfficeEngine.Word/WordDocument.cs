using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeEngine.Diagnostics;
using OfficeEngine.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A WordprocessingML (.docx) document that can be created, opened, edited and saved.
/// </summary>
/// <remarks>
/// <para>
/// A <see cref="WordDocument"/> keeps the underlying OOXML package in memory for its whole
/// lifetime. Parts and elements that this version does not model are preserved when the
/// document is saved; they are reported through <see cref="Diagnostics"/> instead of being
/// dropped silently.
/// </para>
/// <para>
/// All I/O in this version is synchronous. This is a deliberate decision recorded in
/// <c>docs/adr/0004-synchronous-io-and-stream-lifecycle.md</c>. Asynchronous entry points
/// will only be added together with a real asynchronous backing store.
/// </para>
/// <para>
/// This type is thread-affine: instances are not safe for concurrent use.
/// </para>
/// </remarks>
public sealed class WordDocument : IDisposable
{
    private readonly WordprocessingDocument _package;
    private readonly MemoryStream _backing;
    private readonly OfficeEngineDiagnostics _diagnostics = new();
    private bool _disposed;

    private WordDocument(WordprocessingDocument package, MemoryStream backing, OfficeEngineLimits limits)
    {
        _package = package;
        _backing = backing;
        Limits = limits;
        Properties = new DocumentProperties(this);
    }

    /// <summary>Gets the resource limits applied when the document was loaded.</summary>
    public OfficeEngineLimits Limits { get; }

    /// <summary>
    /// Gets the diagnostics reported while the document was loaded.
    /// </summary>
    /// <remarks>
    /// The list is a snapshot; it is not updated by later edits.
    /// </remarks>
    public IReadOnlyList<OfficeEngineDiagnostic> Diagnostics => _diagnostics.ToSnapshot();

    /// <summary>
    /// Gets a live view over the document's core properties (<c>docProps/core.xml</c>).
    /// </summary>
    /// <remarks>
    /// The core properties are the basic metadata of the document: title, subject, author, keywords,
    /// description, last modified by, revision, dates, category, status, identifier, language and
    /// version. Extended properties (<c>docProps/app.xml</c>) and custom properties
    /// (<c>docProps/custom.xml</c>) are not modelled; they are preserved when the document is saved.
    /// </remarks>
    public DocumentProperties Properties { get; }

    /// <summary>Gets a live snapshot of the top-level paragraphs of the main document body.</summary>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public IReadOnlyList<Paragraph> Paragraphs
    {
        get
        {
            ThrowIfDisposed();
            return Body().Elements<W.Paragraph>().Select(element => new Paragraph(element)).ToArray();
        }
    }

    /// <summary>Gets a live snapshot of the tables of the main document body, in document order.</summary>
    /// <remarks>
    /// Only tables that are direct children of the body are returned, in the order they appear — the
    /// same level as <see cref="Paragraphs"/>, because a body interleaves paragraphs and tables.
    /// Tables nested inside a cell are not returned here; this version does not model them, although
    /// it preserves them. Create a table with <see cref="AddTable"/>.
    /// </remarks>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public IReadOnlyList<Table> Tables
    {
        get
        {
            ThrowIfDisposed();
            return Body().Elements<W.Table>().Select(element => new Table(element)).ToArray();
        }
    }

    /// <summary>
    /// Gets a live snapshot of the sections of the document, in document order.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Every document created by <see cref="Create"/> starts with one section: US Letter in portrait
    /// orientation with one-inch margins. The last section is described by the <c>w:sectPr</c> that
    /// closes the body; each earlier section stores its properties in the paragraph that ends it.
    /// Use <see cref="AddSection"/> to append a section.
    /// </para>
    /// <para>
    /// A document loaded from another tool that does not state any section properties reports no
    /// sections here; its content is still valid and may be edited. <see cref="AddSection"/> creates
    /// the missing default section when it is called.
    /// </para>
    /// </remarks>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public IReadOnlyList<Section> Sections
    {
        get
        {
            ThrowIfDisposed();
            MainDocumentPart mainPart = MainPart();
            List<Section> sections = [];
            foreach (OpenXmlElement child in Body().ChildElements)
            {
                if (child is W.Paragraph paragraph)
                {
                    W.SectionProperties? stored = paragraph.ParagraphProperties?.SectionProperties;
                    if (stored is not null)
                    {
                        sections.Add(new Section(stored, mainPart));
                    }
                }
                else if (child is W.SectionProperties section)
                {
                    sections.Add(new Section(section, mainPart));
                }
            }

            return sections;
        }
    }

    /// <summary>
    /// Gets a live snapshot of the style definitions in the document's style part
    /// (<c>word/styles.xml</c>), in document order.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The list is empty when the document has no style part, which every document created by
    /// <see cref="Create"/> is. This version reads the catalogue only: it does not add, change or
    /// remove style definitions, and the style part is preserved exactly as it was loaded.
    /// </para>
    /// <para>
    /// Every definition is exposed, including those whose properties this version does not model, so
    /// a caller can see what a document defines. Resolving how a style is applied — document
    /// defaults, the tables, the direct formatting on a paragraph or a run — is outside this view.
    /// </para>
    /// </remarks>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public IReadOnlyList<WordStyle> Styles
    {
        get
        {
            ThrowIfDisposed();
            W.Styles? root = StyleRoot();
            if (root is null)
            {
                return [];
            }

            return root.Elements<W.Style>().Select(element => new WordStyle(element, root)).ToArray();
        }
    }

    /// <summary>
    /// Finds the style definition with the given identifier, or <see langword="null"/> when the
    /// document does not define it.
    /// </summary>
    /// <remarks>
    /// The identifier is compared with ordinal rules, matching how Office Open XML treats
    /// <c>w:styleId</c>. When a document stores two definitions with the same identifier, which the
    /// schema allows, the first one in document order is returned.
    /// </remarks>
    /// <param name="styleId">The identifier to look up (<c>w:styleId</c>).</param>
    /// <returns>The matching style, or <see langword="null"/> when there is none.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="styleId"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException"><paramref name="styleId"/> is empty.</exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public WordStyle? FindStyle(string styleId)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrEmpty(styleId);

        W.Styles? root = StyleRoot();
        W.Style? element = WordStyle.Find(root, styleId);
        return element is null ? null : new WordStyle(element, root);
    }

    /// <summary>
    /// Gets a live snapshot of the list definitions in the document's numbering part
    /// (<c>word/numbering.xml</c>), in document order.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The list is empty when the document has no numbering part, which every document created by
    /// <see cref="Create"/> is. Each entry is a <c>w:num</c>; its levels are read from the abstract
    /// numbering definition it references. Use <see cref="AddList"/> to create a new one.
    /// </para>
    /// <para>
    /// Every definition is exposed, including those this version does not model, so a caller can see
    /// what a document defines. The numbering part is preserved exactly as it was loaded except for
    /// the definitions <see cref="AddList"/> appends.
    /// </para>
    /// </remarks>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public IReadOnlyList<WordList> Lists
    {
        get
        {
            ThrowIfDisposed();
            W.Numbering? root = NumberingRoot();
            if (root is null)
            {
                return [];
            }

            return root.Elements<W.NumberingInstance>().Select(element => new WordList(element, root)).ToArray();
        }
    }

    /// <summary>
    /// Finds the list definition with the given identifier, or <see langword="null"/> when the
    /// document does not define it.
    /// </summary>
    /// <remarks>
    /// When a document stores two definitions with the same identifier, which the schema allows, the
    /// first one in document order is returned. A definition whose own identifier cannot be read is
    /// skipped rather than failing the lookup.
    /// </remarks>
    /// <param name="id">The identifier to look up (<c>w:num/@w:numId</c>).</param>
    /// <returns>The matching list, or <see langword="null"/> when there is none.</returns>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public WordList? FindList(int id)
    {
        ThrowIfDisposed();

        W.Numbering? root = NumberingRoot();
        W.NumberingInstance? element = WordList.Find(root, id);
        return element is null ? null : new WordList(element, root);
    }

    /// <summary>
    /// Creates a new list definition with a single level and returns it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The list is added to the document's numbering part, creating that part when the document does
    /// not have one. The new definition is immediately usable: apply it to a paragraph with
    /// <see cref="Paragraph.ApplyList"/>, or assign its <see cref="WordList.Id"/> to
    /// <see cref="Paragraph.ListId"/>.
    /// </para>
    /// <para>
    /// The single level is numbered with the requested format and laid out with the usual half-inch
    /// hanging indent. A bullet level uses the round bullet and the <c>Symbol</c> font that Word
    /// writes for its own bulleted lists.
    /// </para>
    /// </remarks>
    /// <param name="format">The numbering format of the level.</param>
    /// <returns>The new list definition.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is not a defined format.</exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public WordList AddList(WordListFormat format)
    {
        ThrowIfDisposed();

        // Map before touching the part so that a rejected format leaves the document unchanged.
        W.NumberFormatValues numberFormat = ToOoxmlListFormat(format);

        MainDocumentPart mainPart = _package.MainDocumentPart
            ?? throw new InvalidOfficeFileException("The document no longer has a main document part.");
        NumberingDefinitionsPart part = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering ??= new W.Numbering();
        W.Numbering root = part.Numbering;

        int abstractNumberId = NextAbstractNumberId(root);
        int listId = NextListId(root);

        W.AbstractNum abstractNumbering = BuildAbstractNumbering(numberFormat, abstractNumberId);
        W.NumberingInstance instance = new(new W.AbstractNumId { Val = abstractNumberId }) { NumberID = listId };

        // CT_Numbering requires every w:abstractNum (and w:numPicBullet) to precede every w:num.
        W.NumberingInstance? firstInstance = root.Elements<W.NumberingInstance>().FirstOrDefault();
        if (firstInstance is null)
        {
            root.AppendChild(abstractNumbering);
        }
        else
        {
            root.InsertBefore(abstractNumbering, firstInstance);
        }

        root.AppendChild(instance);
        return new WordList(instance, root);
    }

    /// <summary>
    /// Gets the visible text of the document: one line per top-level paragraph and one line per table
    /// row, joined by line-feed characters.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Each top-level paragraph contributes its own text (<see cref="Paragraph.Text"/>). Each table
    /// contributes one line per row, with the text of its cells separated by tab characters; a cell
    /// that holds several paragraphs contributes them on separate lines.
    /// </para>
    /// <para>
    /// Only the body is read, and only the content this version models: text inside a nested table,
    /// and text carried by structures that are not modelled (for example a content control), is not
    /// included. It is preserved when the document is saved.
    /// </para>
    /// </remarks>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public string Text
    {
        get
        {
            ThrowIfDisposed();
            return BlockText.Of(Body().ChildElements);
        }
    }

    /// <summary>
    /// Finds the occurrences of a text in the body of the document.
    /// </summary>
    /// <param name="text">The text to look for. It cannot be empty.</param>
    /// <param name="matchCase">Whether the comparison distinguishes upper and lower case.</param>
    /// <returns>The occurrences, in document order.</returns>
    /// <remarks>
    /// <para>
    /// The scope is the body: the top-level paragraphs and the paragraphs of the top-level tables.
    /// Headers, footers, notes, comments and the paragraphs of nested tables are not searched, and
    /// the search never crosses the boundary of a paragraph.
    /// </para>
    /// <para>
    /// The searched text is the one of the direct runs of each paragraph and of the runs of its direct
    /// hyperlinks.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentException"><paramref name="text"/> is empty.</exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public IReadOnlyList<WordTextMatch> FindText(string text, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrEmpty(text);
        List<WordTextMatch> matches = [];
        foreach (Paragraph paragraph in SearchableParagraphs())
        {
            matches.AddRange(paragraph.FindText(text, matchCase));
        }

        return matches;
    }

    /// <summary>Replaces the occurrences of a text in the body of the document.</summary>
    /// <param name="oldText">The text to replace. It cannot be empty.</param>
    /// <param name="newText">The text that replaces it. An empty text removes the match.</param>
    /// <param name="matchCase">Whether the comparison distinguishes upper and lower case.</param>
    /// <returns>The number of replacements.</returns>
    /// <remarks>
    /// The scope is the body: the top-level paragraphs and the paragraphs of the top-level tables.
    /// The replaced text keeps the character formatting of the first run of the match; see
    /// <see cref="Paragraph.ReplaceText"/>.
    /// </remarks>
    /// <exception cref="ArgumentException"><paramref name="oldText"/> is empty.</exception>
    /// <exception cref="ArgumentNullException"><paramref name="newText"/> is <see langword="null"/>.</exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public int ReplaceText(string oldText, string newText, bool matchCase = false)
    {
        ArgumentException.ThrowIfNullOrEmpty(oldText);
        ArgumentNullException.ThrowIfNull(newText);
        int replaced = 0;
        foreach (Paragraph paragraph in SearchableParagraphs())
        {
            replaced += paragraph.ReplaceText(oldText, newText, matchCase);
        }

        return replaced;
    }

    /// <summary>Creates a new, empty document in memory.</summary>
    /// <remarks>
    /// The new document has a single section: US Letter in portrait orientation with one-inch
    /// margins. Change it through <see cref="Sections"/>.
    /// </remarks>
    /// <param name="limits">Optional resource limits. Defaults to <see cref="OfficeEngineLimits.Default"/>.</param>
    /// <returns>A new document.</returns>
    public static WordDocument Create(OfficeEngineLimits? limits = null)
    {
        OfficeEngineLimits effective = limits ?? OfficeEngineLimits.Default;
        effective.EnsureValid();

        MemoryStream backing = new();
        WordprocessingDocument? package = null;
        try
        {
            package = WordprocessingDocument.Create(backing, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = package.AddMainDocumentPart();
            mainPart.Document = new W.Document(new W.Body(Section.CreateDefault()));
            return new WordDocument(package, backing, effective);
        }
        catch
        {
            package?.Dispose();
            backing.Dispose();
            throw;
        }
    }

    /// <summary>Opens a document from a file.</summary>
    /// <param name="path">Path of the .docx file.</param>
    /// <param name="limits">Optional resource limits. Defaults to <see cref="OfficeEngineLimits.Default"/>.</param>
    /// <returns>The opened document.</returns>
    /// <exception cref="InvalidOfficeFileException">The file is not a readable WordprocessingML document.</exception>
    public static WordDocument Open(string path, OfficeEngineLimits? limits = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        OfficeEngineLimits effective = limits ?? OfficeEngineLimits.Default;
        MemoryStream backing = OoxmlLoader.LoadFromFile(path, effective);
        return OpenCore(backing, effective);
    }

    /// <summary>Opens a document from a stream.</summary>
    /// <param name="stream">The stream that contains the .docx package.</param>
    /// <param name="leaveOpen">
    /// When <see langword="false"/> (the default), <paramref name="stream"/> is disposed after it has been read.
    /// The document never holds a reference to the passed stream.
    /// </param>
    /// <param name="limits">Optional resource limits. Defaults to <see cref="OfficeEngineLimits.Default"/>.</param>
    /// <returns>The opened document.</returns>
    /// <exception cref="InvalidOfficeFileException">The stream is not a readable WordprocessingML document.</exception>
    public static WordDocument Open(Stream stream, bool leaveOpen = false, OfficeEngineLimits? limits = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        OfficeEngineLimits effective = limits ?? OfficeEngineLimits.Default;
        MemoryStream backing = OoxmlLoader.LoadFromStream(stream, leaveOpen, effective);
        return OpenCore(backing, effective);
    }

    /// <summary>Adds a new paragraph with the given text at the end of the document body.</summary>
    /// <param name="text">The text of the paragraph. May be empty.</param>
    /// <returns>A live view over the new paragraph.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="text"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store. Add tabs and breaks with <see cref="Run.AddTab"/> and <see cref="Run.AddBreak"/>.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public Paragraph AddParagraph(string text)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(text);
        PlainText.EnsureValid(text, nameof(text));

        W.Paragraph element = new();
        if (text.Length > 0)
        {
            element.AppendChild(new W.Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        AppendBodyBlock(element);
        return new Paragraph(element);
    }

    /// <summary>Adds a new table with the given number of rows and columns at the end of the document body.</summary>
    /// <remarks>
    /// <para>
    /// Every cell starts with a single empty paragraph, which Word needs to edit a cell. Columns are
    /// fixed when the table is created: <see cref="Table.AddRow"/> grows a table, but this version does
    /// not add or remove columns. Table and cell formatting (borders, shading, widths, merge) is not
    /// modelled yet.
    /// </para>
    /// <para>
    /// The column widths written into the grid are a fixed default rather than a measurement of the
    /// page, because the table grid is not modelled in this version.
    /// </para>
    /// </remarks>
    /// <param name="rows">The number of rows. Must be at least one.</param>
    /// <param name="columns">The number of columns. Must be at least one.</param>
    /// <returns>A live view over the new table.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="rows"/> or <paramref name="columns"/> is less than one.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public Table AddTable(int rows, int columns)
    {
        ThrowIfDisposed();
        if (rows < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(rows), rows, "A table must have at least one row.");
        }

        if (columns < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(columns), columns, "A table must have at least one column.");
        }

        W.Table element = Table.Create(rows, columns);
        AppendBodyBlock(element);
        return new Table(element);
    }

    /// <summary>
    /// Appends a new section to the end of the document and returns it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The section that was last is closed with a paragraph that carries its properties, and the body
    /// keeps a final <c>w:sectPr</c> for the new section. The new section starts with the same page
    /// size, orientation and margins as the previous one, so they can be changed independently
    /// afterwards through <see cref="Sections"/>.
    /// </para>
    /// <para>
    /// The section break is a paragraph: it appears in <see cref="Paragraphs"/> as an empty paragraph,
    /// which is also how Word stores a section break. A document loaded without section properties
    /// gets a single default section (US Letter, portrait, one-inch margins) and no break paragraph.
    /// </para>
    /// <para>
    /// A <see cref="Section"/> obtained before this call keeps describing the section that just
    /// ended, so it can still be edited safely afterwards.
    /// </para>
    /// </remarks>
    /// <param name="startType">How the new section starts. Defaults to a new page.</param>
    /// <returns>The new section, which is the last one.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="startType"/> is not a defined value.</exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public Section AddSection(SectionStartType startType = SectionStartType.NextPage)
    {
        ThrowIfDisposed();

        // Validate before touching the tree so that a rejected value leaves the document unchanged.
        W.SectionMarkValues mark = Section.ToOoxmlStartType(startType);

        MainDocumentPart mainPart = MainPart();
        W.Body body = Body();
        W.SectionProperties? final = body.Elements<W.SectionProperties>().LastOrDefault();
        if (final is null)
        {
            final = Section.CreateDefault();
            body.AppendChild(final);
            return new Section(final, mainPart);
        }

        // The element that was the body's final w:sectPr becomes the properties of the section that
        // just ended, so wrappers created before this call keep pointing at that same section. The
        // body gets a clone for the new last section.
        W.SectionProperties carried = final;
        W.SectionProperties next = (W.SectionProperties)final.CloneNode(deep: true);
        OpenXmlElement? marker = final.NextSibling();
        body.RemoveChild(final);

        W.Paragraph breakParagraph = new(new W.ParagraphProperties(carried));
        if (marker is null)
        {
            body.AppendChild(breakParagraph);
            body.AppendChild(next);
        }
        else
        {
            body.InsertBefore(breakParagraph, marker);
            body.InsertBefore(next, marker);
        }

        next.RemoveAllChildren<W.SectionType>();
        if (startType != SectionStartType.NextPage)
        {
            Section.WriteStartType(next, mark);
        }

        return new Section(next, mainPart);
    }

    /// <summary>Saves the document to a file, replacing any existing file atomically.</summary>
    /// <param name="path">Destination path.</param>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public void Save(string path)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrEmpty(path);

        string fullPath = Path.GetFullPath(path);
        string directory = Path.GetDirectoryName(fullPath) ?? ".";
        string temporaryPath = Path.Combine(directory, $"{Path.GetRandomFileName()}.tmp");
        try
        {
            using (FileStream file = new(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                Save(file);
            }

            File.Move(temporaryPath, fullPath, overwrite: true);
        }
        catch
        {
            TryDelete(temporaryPath);
            throw;
        }
    }

    /// <summary>Saves the document to a stream.</summary>
    /// <param name="destination">The writable destination stream. It is left open.</param>
    /// <exception cref="ArgumentNullException"><paramref name="destination"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="ObjectDisposedException">The document has been disposed.</exception>
    public void Save(Stream destination)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(destination);
        if (!destination.CanWrite)
        {
            throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        }

        _package.Save();
        byte[] bytes = _backing.ToArray();
        destination.Write(bytes, 0, bytes.Length);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _package.Dispose();
        _backing.Dispose();
    }

    private IEnumerable<Paragraph> SearchableParagraphs()
    {
        foreach (W.Paragraph element in Body().Elements<W.Paragraph>())
        {
            yield return new Paragraph(element);
        }

        foreach (W.Table table in Body().Elements<W.Table>())
        {
            foreach (W.TableRow row in table.Elements<W.TableRow>())
            {
                foreach (W.TableCell cell in row.Elements<W.TableCell>())
                {
                    foreach (W.Paragraph element in cell.Elements<W.Paragraph>())
                    {
                        yield return new Paragraph(element);
                    }
                }
            }
        }
    }

    private static WordDocument OpenCore(MemoryStream backing, OfficeEngineLimits limits)
    {
        WordprocessingDocument? package = null;
        try
        {
            OoxmlPackageInspector.EnsureWithinLimits(backing, limits);
            OoxmlXmlGuard.EnsureSafe(backing, limits);
            backing.Position = 0;

            try
            {
                package = WordprocessingDocument.Open(backing, isEditable: true);
                MainDocumentPart mainPart = package.MainDocumentPart
                    ?? throw new InvalidOfficeFileException("The package does not contain a main document part, so it is not a WordprocessingML document.");
                mainPart.Document ??= new W.Document(new W.Body());
                mainPart.Document.Body ??= new W.Body();
            }
            catch (Exception ex) when (ex is OpenXmlPackageException or FileFormatException
                or InvalidDataException or InvalidOperationException or XmlException or ArgumentOutOfRangeException)
            {
                throw new InvalidOfficeFileException("The input is not a readable WordprocessingML (.docx) document.", ex);
            }

            WordDocument document = new(package, backing, limits);
            document.Inspect();
            return document;
        }
        catch
        {
            package?.Dispose();
            backing.Dispose();
            throw;
        }
    }

    private void Inspect()
    {
        try
        {
            InspectCore();
        }
        catch (Exception ex) when (ex is InvalidDataException or XmlException or OpenXmlPackageException
            or FileFormatException or InvalidOperationException or ArgumentOutOfRangeException)
        {
            throw new InvalidOfficeFileException("The package cannot be read as a WordprocessingML (.docx) document.", ex);
        }
    }

    private void InspectCore()
    {
        W.Body body = Body();
        int unmodelled = 0;
        List<string> elementNames = [];
        foreach (OpenXmlElement child in body.ChildElements)
        {
            if (child is W.Paragraph or W.Table or W.SectionProperties)
            {
                continue;
            }

            unmodelled++;
            if (!elementNames.Contains(child.LocalName))
            {
                elementNames.Add(child.LocalName);
            }
        }

        if (unmodelled > 0)
        {
            _diagnostics.Add(
                DiagnosticSeverity.Info,
                WordDiagnosticCodes.UnmodelledBodyContent,
                $"The document body contains {unmodelled} top-level element(s) that this version does not model " +
                $"({string.Join(", ", elementNames)}). They are preserved when the document is saved but cannot be edited through the API.");
        }
    }

    internal WordprocessingDocument Package
    {
        get
        {
            ThrowIfDisposed();
            return _package;
        }
    }

    private MainDocumentPart MainPart()
        => _package.MainDocumentPart
            ?? throw new InvalidOfficeFileException("The document no longer has a main document part.");

    private W.Body Body()
    {
        W.Document document = OoxmlPartReader.Read("the main document part", () => MainPart().Document);
        return document.Body
            ?? throw new InvalidOfficeFileException("The document has no body element.");
    }

    /// <summary>
    /// Appends a block-level element to the body, before the final section properties when the
    /// document has them: <c>w:sectPr</c> must be the last child of <c>w:body</c>.
    /// </summary>
    private void AppendBodyBlock(OpenXmlElement element)
    {
        W.Body body = Body();
        W.SectionProperties? final = body.Elements<W.SectionProperties>().LastOrDefault();
        if (final is null)
        {
            body.AppendChild(element);
        }
        else
        {
            body.InsertBefore(element, final);
        }
    }

    /// <summary>
    /// Gets the root of the document's style part, or <see langword="null"/> when the document has
    /// no style part.
    /// </summary>
    /// <remarks>
    /// A style part that exists but has no <c>w:styles</c> root is reported the same way as a missing
    /// style part: there is no catalogue to read.
    /// </remarks>
    private W.Styles? StyleRoot()
    {
        MainDocumentPart mainPart = _package.MainDocumentPart
            ?? throw new InvalidOfficeFileException("The document no longer has a main document part.");
        return OoxmlPartReader.ReadOrDefault("the styles part", () => mainPart.StyleDefinitionsPart?.Styles);
    }

    /// <summary>
    /// Gets the root of the document's numbering part, or <see langword="null"/> when the document
    /// has no numbering part or it has no <c>w:numbering</c> root.
    /// </summary>
    private W.Numbering? NumberingRoot()
    {
        MainDocumentPart mainPart = _package.MainDocumentPart
            ?? throw new InvalidOfficeFileException("The document no longer has a main document part.");
        return OoxmlPartReader.ReadOrDefault("the numbering part", () => mainPart.NumberingDefinitionsPart?.Numbering);
    }

    private static int NextAbstractNumberId(W.Numbering root)
    {
        int highest = -1;
        foreach (W.AbstractNum definition in root.Elements<W.AbstractNum>())
        {
            if (ListReference.TryReadInt(definition.AbstractNumberId) is { } id && id > highest)
            {
                highest = id;
            }
        }

        return highest + 1;
    }

    private static int NextListId(W.Numbering root)
    {
        // 0 is reserved for "the paragraph is not numbered", so the first generated id is 1.
        int highest = 0;
        foreach (W.NumberingInstance instance in root.Elements<W.NumberingInstance>())
        {
            if (ListReference.TryReadInt(instance.NumberID) is { } id && id > highest)
            {
                highest = id;
            }
        }

        return highest + 1;
    }

    private static W.AbstractNum BuildAbstractNumbering(W.NumberFormatValues format, int abstractNumberId)
    {
        bool isBullet = format == W.NumberFormatValues.Bullet;
        W.Level level = new()
        {
            LevelIndex = 0,
            StartNumberingValue = new W.StartNumberingValue { Val = 1 },
            NumberingFormat = new W.NumberingFormat { Val = format },
            LevelText = new W.LevelText { Val = isBullet ? "\uF0B7" : "%1." },
            LevelJustification = new W.LevelJustification { Val = W.LevelJustificationValues.Left },
            PreviousParagraphProperties = new W.PreviousParagraphProperties(
                new W.Indentation { Left = "720", Hanging = "360" }),
        };

        if (isBullet)
        {
            level.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties(
                new W.RunFonts { Ascii = "Symbol", HighAnsi = "Symbol", Hint = W.FontTypeHintValues.Default });
        }

        return new W.AbstractNum(level)
        {
            AbstractNumberId = abstractNumberId,
            MultiLevelType = new W.MultiLevelType { Val = W.MultiLevelValues.SingleLevel },
        };
    }

    private static W.NumberFormatValues ToOoxmlListFormat(WordListFormat value)
    {
        if (value == WordListFormat.Bullet)
        {
            return W.NumberFormatValues.Bullet;
        }

        if (value == WordListFormat.Decimal)
        {
            return W.NumberFormatValues.Decimal;
        }

        if (value == WordListFormat.LowerLetter)
        {
            return W.NumberFormatValues.LowerLetter;
        }

        if (value == WordListFormat.UpperLetter)
        {
            return W.NumberFormatValues.UpperLetter;
        }

        if (value == WordListFormat.LowerRoman)
        {
            return W.NumberFormatValues.LowerRoman;
        }

        if (value == WordListFormat.UpperRoman)
        {
            return W.NumberFormatValues.UpperRoman;
        }

        throw new ArgumentOutOfRangeException(nameof(value), value, "The list format is not a defined value.");
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // Best effort cleanup; the original exception is more important.
        }
        catch (UnauthorizedAccessException)
        {
            // Best effort cleanup; the original exception is more important.
        }
    }

    private void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);
}
