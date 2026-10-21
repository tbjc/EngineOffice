using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over an external hyperlink in a <see cref="WordDocument"/>.
/// </summary>
/// <remarks>
/// <para>
/// An external hyperlink is a <c>w:hyperlink</c> whose target lives in a relationship
/// (<c>r:id</c>) owned by the part that contains the paragraph: <c>word/document.xml</c>, a header
/// or a footer. This wrapper exposes the visible text and the destination; the relationship
/// itself stays private.
/// </para>
/// <para>
/// Internal links to a bookmark (<c>w:anchor</c>), links without a relationship and links nested in
/// structures this version does not model are preserved when the document is saved but are not
/// returned by <see cref="Paragraph.Hyperlinks"/> and cannot be edited through this API.
/// </para>
/// <para>
/// This is a lightweight wrapper: it does not own the underlying content. Reading or writing it
/// mutates the document that produced it, and handles must not be used after that document has
/// been disposed.
/// </para>
/// </remarks>
public sealed class Hyperlink
{
    private readonly W.Hyperlink _element;

    internal Hyperlink(W.Hyperlink element) => _element = element;

    /// <summary>
    /// Gets or sets the visible text of the hyperlink.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The getter concatenates every <c>w:t</c> text element inside the hyperlink, including text
    /// nested in structures this version does not model. Tabs, line breaks and drawings contribute
    /// nothing to the string.
    /// </para>
    /// <para>
    /// The setter replaces the hyperlink's content with a single run carrying the new text. The
    /// destination and the hyperlink's own attributes are preserved; other content the hyperlink
    /// carried, such as bookmarks or character formatting, is removed with it. Use
    /// <see cref="Runs"/> to edit the text of an existing run without losing its formatting.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="value"/> contains a tab, a line break or a character that Office Open XML
    /// cannot store. Add tabs and breaks through <see cref="Run.AddTab"/> and
    /// <see cref="Run.AddBreak"/>.
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

            foreach (OpenXmlElement child in _element.ChildElements.ToArray())
            {
                child.Remove();
            }

            if (value.Length > 0)
            {
                _element.AppendChild(new W.Run(new W.Text(value) { Space = SpaceProcessingModeValues.Preserve }));
            }
        }
    }

    /// <summary>
    /// Gets a live view over every run that is a direct child of this hyperlink.
    /// </summary>
    /// <remarks>
    /// The list reflects the current state of the hyperlink and is rebuilt on each access. Runs
    /// nested inside structures this version does not model are not returned. Editing a run keeps
    /// its character formatting, unlike assigning <see cref="Text"/>.
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
    /// Gets or sets the external destination of the hyperlink.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The value is stored in the hyperlink relationship of the owning part, so the package may
    /// normalize it (for example, by adding a trailing slash to an absolute HTTP URI). The API does
    /// not validate the URI scheme; OfficeEngine stores the destination but never opens or executes
    /// it.
    /// </para>
    /// <para>
    /// Assigning a new value creates a new relationship for this hyperlink. When the previous
    /// relationship is not used by another hyperlink in the same part, it is removed so the package
    /// does not accumulate orphan relationships. A relationship shared with another hyperlink is
    /// left in place, so editing one link does not change the others.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentNullException">The assigned value is <see langword="null"/>.</exception>
    /// <exception cref="InvalidOfficeFileException">
    /// The hyperlink has no relationship id, or the relationship id does not resolve to an external
    /// hyperlink relationship.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The owning document has been disposed.</exception>
    public Uri Destination
    {
        get => GetRelationship(_element).Uri;
        set
        {
            ArgumentNullException.ThrowIfNull(value);

            OpenXmlPart owner = ResolvePart(_element);
            string? previousId = _element.Id?.Value;
            HyperlinkRelationship? previous = previousId is null
                ? null
                : owner.HyperlinkRelationships.FirstOrDefault(relationship => relationship.Id == previousId);

            if (previous is not null && previous.IsExternal && previous.Uri.Equals(value))
            {
                return;
            }

            HyperlinkRelationship created = AddRelationship(owner, value);
            _element.Id = created.Id;

            if (previousId is not null)
            {
                DeleteIfUnused(owner, previousId);
            }
        }
    }

    /// <inheritdoc />
    public override string ToString() => Text;

    /// <summary>
    /// Resolves the external relationship a stored hyperlink points at, failing with a typed error
    /// when the stored package is inconsistent.
    /// </summary>
    internal static HyperlinkRelationship GetRelationship(W.Hyperlink element)
    {
        StringValue? storedId = element.Id;
        if (storedId is null)
        {
            throw new InvalidOfficeFileException("A hyperlink stores no relationship id.");
        }

        string relationshipId = storedId.Value
            ?? throw new InvalidOfficeFileException("A hyperlink stores an empty relationship id.");

        OpenXmlPart owner = ResolvePart(element);
        HyperlinkRelationship? relationship = owner.HyperlinkRelationships
            .FirstOrDefault(candidate => candidate.Id == relationshipId);

        if (relationship is null)
        {
            throw new InvalidOfficeFileException(
                $"A hyperlink references a relationship ('{relationshipId}') that is missing or is not a hyperlink relationship.");
        }

        if (!relationship.IsExternal)
        {
            throw new InvalidOfficeFileException(
                $"A hyperlink references a relationship ('{relationshipId}') that is not external.");
        }

        return relationship;
    }

    /// <summary>Creates a hyperlink with an external relationship and appends it to a paragraph.</summary>
    internal static Hyperlink Add(W.Paragraph paragraph, string text, Uri destination)
    {
        OpenXmlPart owner = ResolvePart(paragraph);
        HyperlinkRelationship relationship = AddRelationship(owner, destination);
        W.Hyperlink element = new() { Id = relationship.Id };
        if (text.Length > 0)
        {
            element.AppendChild(new W.Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        paragraph.AppendChild(element);
        return new Hyperlink(element);
    }

    private static OpenXmlPart ResolvePart(OpenXmlElement element)
    {
        OpenXmlPartRootElement? root = element.Ancestors<OpenXmlPartRootElement>().FirstOrDefault();
        return root?.OpenXmlPart
            ?? throw new InvalidOfficeFileException("The hyperlink is not attached to a document part.");
    }

    private static HyperlinkRelationship AddRelationship(OpenXmlPart owner, Uri destination) => owner switch
    {
        MainDocumentPart main => main.AddHyperlinkRelationship(destination, isExternal: true),
        HeaderPart header => header.AddHyperlinkRelationship(destination, isExternal: true),
        FooterPart footer => footer.AddHyperlinkRelationship(destination, isExternal: true),
        _ => throw new InvalidOperationException(
            $"Hyperlinks cannot be added to a '{owner.GetType().Name}' part."),
    };

    private static void DeleteIfUnused(OpenXmlPart owner, string relationshipId)
    {
        OpenXmlPartRootElement? root = owner.RootElement;
        if (root is not null
            && root.Descendants<W.Hyperlink>().Any(link => link.Id?.Value == relationshipId))
        {
            return;
        }

        HyperlinkRelationship? relationship = owner.HyperlinkRelationships
            .FirstOrDefault(candidate => candidate.Id == relationshipId);
        if (relationship is not null)
        {
            owner.DeleteReferenceRelationship(relationship);
        }
    }
}
