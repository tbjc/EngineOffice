using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A read-only view over a single list definition (<c>w:num</c>) of a <see cref="WordDocument"/>.
/// </summary>
/// <remarks>
/// <para>
/// A list definition lives in the document's numbering part (<c>word/numbering.xml</c>). A paragraph
/// joins a list by carrying its identifier in <c>w:numPr/w:numId</c> (see
/// <see cref="Paragraph.ListId"/>). Several paragraphs that share the identifier share the list, and
/// the list's numbering continues across them.
/// </para>
/// <para>
/// A <c>w:num</c> carries little itself: it references an abstract numbering definition
/// (<c>w:abstractNum</c>) through <c>w:abstractNumId</c>, and that definition holds the levels.
/// <see cref="Levels"/> resolves that reference. When the identifier points to a definition the part
/// does not contain, which Office Open XML allows, the list has no levels and Word renders its
/// paragraphs as if they were not numbered.
/// </para>
/// <para>
/// New lists are created with <see cref="WordDocument.AddList"/>. This version does not otherwise
/// add, change or remove definitions, and the numbering part is preserved exactly as it was loaded.
/// </para>
/// <para>
/// This is a lightweight wrapper: it does not own the underlying content and must not be used after
/// the <see cref="WordDocument"/> that produced it has been disposed.
/// </para>
/// </remarks>
public sealed class WordList
{
    private readonly W.NumberingInstance _element;
    private readonly W.Numbering? _root;

    internal WordList(W.NumberingInstance element, W.Numbering? root)
    {
        _element = element;
        _root = root;
    }

    /// <summary>
    /// Gets the identifier of the list (<c>w:num/@w:numId</c>), the value a paragraph stores in
    /// <c>w:numPr/w:numId</c> to join it.
    /// </summary>
    /// <exception cref="InvalidOfficeFileException">
    /// The list has no identifier or the stored value cannot be parsed.
    /// </exception>
    public int Id
    {
        get
        {
            Int32Value? stored = _element.NumberID;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A list stores no w:numId attribute, which the schema requires.");
            }

            return ReadStoredValue(() => stored.Value, "list id");
        }
    }

    /// <summary>
    /// Gets the identifier of the abstract numbering definition this list references
    /// (<c>w:abstractNumId</c>).
    /// </summary>
    /// <remarks>
    /// The identifier is not resolved here: it may point to a definition the part does not contain,
    /// in which case <see cref="Levels"/> is empty. Use <see cref="Levels"/> or
    /// <see cref="FindLevel"/> to read the levels the identifier leads to.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">
    /// The list has no abstract numbering reference, it has one without a value, or the stored value
    /// cannot be parsed.
    /// </exception>
    public int AbstractNumberId
    {
        get
        {
            W.AbstractNumId? abstractNumberId = _element.GetFirstChild<W.AbstractNumId>();
            if (abstractNumberId is null)
            {
                throw new InvalidOfficeFileException("A list stores no w:abstractNumId element, which the schema requires.");
            }

            Int32Value? stored = abstractNumberId.Val;
            if (stored is null)
            {
                throw new InvalidOfficeFileException("A list stores a w:abstractNumId element without the required w:val attribute.");
            }

            return ReadStoredValue(() => stored.Value, "abstract numbering id");
        }
    }

    /// <summary>
    /// Gets the levels of the abstract numbering definition this list references, in document order.
    /// </summary>
    /// <remarks>
    /// The list is empty when the referenced definition is not present in the numbering part, which
    /// Office Open XML allows and Word ignores.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">
    /// The list's abstract numbering reference cannot be read.
    /// </exception>
    public IReadOnlyList<WordListLevel> Levels
    {
        get
        {
            W.AbstractNum? definition = FindAbstract(_root, AbstractNumberId);
            if (definition is null)
            {
                return [];
            }

            return definition.Elements<W.Level>().Select(element => new WordListLevel(element)).ToArray();
        }
    }

    /// <summary>
    /// Finds the level with the given index, or <see langword="null"/> when the list does not define
    /// one.
    /// </summary>
    /// <param name="level">The level index to look up, from <c>0</c> to <c>8</c>.</param>
    /// <returns>The matching level, or <see langword="null"/> when there is none.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="level"/> is negative or greater than <see cref="ListReference.MaxLevel"/>.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">
    /// The list's abstract numbering reference cannot be read.
    /// </exception>
    public WordListLevel? FindLevel(int level)
    {
        ListReference.EnsureValidLevel(level, nameof(level));

        W.Level? element = WordListLevel.Find(FindAbstract(_root, AbstractNumberId), level);
        return element is null ? null : new WordListLevel(element);
    }

    /// <inheritdoc />
    public override string ToString()
    {
        // The raw text is reported so that ToString never throws, even for a list whose identifier
        // cannot be parsed.
        return _element.NumberID?.InnerText is { Length: > 0 } id ? id : "(list without an id)";
    }

    /// <summary>
    /// Finds the first <c>w:num</c> in <paramref name="root"/> whose identifier equals
    /// <paramref name="id"/>.
    /// </summary>
    internal static W.NumberingInstance? Find(W.Numbering? root, int id)
    {
        if (root is null)
        {
            return null;
        }

        foreach (W.NumberingInstance candidate in root.Elements<W.NumberingInstance>())
        {
            if (ListReference.TryReadInt(candidate.NumberID) == id)
            {
                return candidate;
            }
        }

        return null;
    }

    private static W.AbstractNum? FindAbstract(W.Numbering? root, int abstractNumberId)
    {
        if (root is null)
        {
            return null;
        }

        foreach (W.AbstractNum candidate in root.Elements<W.AbstractNum>())
        {
            if (ListReference.TryReadInt(candidate.AbstractNumberId) == abstractNumberId)
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
                $"A list stores a {propertyDescription} value that this document cannot read.",
                exception);
        }
    }
}
