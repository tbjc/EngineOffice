using System;
using System.IO;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over the core properties of a <see cref="WordDocument"/>.
/// </summary>
/// <remarks>
/// <para>
/// The core properties are the basic metadata stored in <c>docProps/core.xml</c>: title, subject,
/// author, keywords, description, last modified by, revision, dates, category, status, identifier,
/// language and version. Reading a property the document does not store returns
/// <see langword="null"/>; assigning <see langword="null"/> removes it again.
/// </para>
/// <para>
/// Extended properties (<c>docProps/app.xml</c>) and custom properties
/// (<c>docProps/custom.xml</c>) are not modelled here. They are preserved when the document is
/// saved. The core properties part is created only when a value is set; reading a document without
/// one does not add an empty part.
/// </para>
/// <para>
/// String values may contain tabs and line breaks, but a character XML cannot store is rejected
/// with <see cref="ArgumentException"/>. Dates are stored as an instant with a time-zone offset; the
/// SDK reports them as local <see cref="DateTime"/> values, so compare instants through
/// <see cref="DateTime.ToUniversalTime"/> when the original offset matters.
/// </para>
/// <para>
/// This is a live view: it mutates the document that produced it and must not be used after that
/// document has been disposed. Every member throws <see cref="ObjectDisposedException"/> in that
/// case.
/// </para>
/// </remarks>
public sealed class DocumentProperties
{
    private const string EmptyCorePropertiesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
        "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" " +
        "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>";

    private readonly WordDocument _document;

    internal DocumentProperties(WordDocument document) => _document = document;

    /// <summary>Gets or sets the title (<c>dc:title</c>).</summary>
    public string? Title
    {
        get => Read(static part => part.CoreFileProperties.Title);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Title = text);
    }

    /// <summary>Gets or sets the subject (<c>dc:subject</c>).</summary>
    public string? Subject
    {
        get => Read(static part => part.CoreFileProperties.Subject);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Subject = text);
    }

    /// <summary>Gets or sets the author (<c>dc:creator</c>).</summary>
    public string? Creator
    {
        get => Read(static part => part.CoreFileProperties.Creator);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Creator = text);
    }

    /// <summary>Gets or sets the keywords (<c>cp:keywords</c>).</summary>
    public string? Keywords
    {
        get => Read(static part => part.CoreFileProperties.Keywords);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Keywords = text);
    }

    /// <summary>Gets or sets the description or comments (<c>dc:description</c>).</summary>
    public string? Description
    {
        get => Read(static part => part.CoreFileProperties.Description);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Description = text);
    }

    /// <summary>Gets or sets the last editor (<c>cp:lastModifiedBy</c>).</summary>
    public string? LastModifiedBy
    {
        get => Read(static part => part.CoreFileProperties.LastModifiedBy);
        set => SetString(value, static (part, text) => part.CoreFileProperties.LastModifiedBy = text);
    }

    /// <summary>Gets or sets the revision (<c>cp:revision</c>).</summary>
    /// <remarks>
    /// The value is stored as text, not as a number, because that is how Office Open XML defines it.
    /// </remarks>
    public string? Revision
    {
        get => Read(static part => part.CoreFileProperties.Revision);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Revision = text);
    }

    /// <summary>Gets or sets the last print instant (<c>cp:lastPrinted</c>).</summary>
    public DateTime? LastPrinted
    {
        get => Read(static part => part.CoreFileProperties.LastPrinted);
        set => SetDate(value, static (part, date) => part.CoreFileProperties.LastPrinted = date);
    }

    /// <summary>Gets or sets the creation instant (<c>dcterms:created</c>).</summary>
    public DateTime? Created
    {
        get => Read(static part => part.CoreFileProperties.Created);
        set => SetDate(value, static (part, date) => part.CoreFileProperties.Created = date);
    }

    /// <summary>Gets or sets the last modification instant (<c>dcterms:modified</c>).</summary>
    public DateTime? Modified
    {
        get => Read(static part => part.CoreFileProperties.Modified);
        set => SetDate(value, static (part, date) => part.CoreFileProperties.Modified = date);
    }

    /// <summary>Gets or sets the category (<c>cp:category</c>).</summary>
    public string? Category
    {
        get => Read(static part => part.CoreFileProperties.Category);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Category = text);
    }

    /// <summary>Gets or sets the content status (<c>cp:contentStatus</c>).</summary>
    public string? ContentStatus
    {
        get => Read(static part => part.CoreFileProperties.ContentStatus);
        set => SetString(value, static (part, text) => part.CoreFileProperties.ContentStatus = text);
    }

    /// <summary>Gets or sets the identifier (<c>dc:identifier</c>).</summary>
    public string? Identifier
    {
        get => Read(static part => part.CoreFileProperties.Identifier);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Identifier = text);
    }

    /// <summary>Gets or sets the language (<c>dc:language</c>).</summary>
    public string? Language
    {
        get => Read(static part => part.CoreFileProperties.Language);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Language = text);
    }

    /// <summary>Gets or sets the version (<c>cp:version</c>).</summary>
    public string? Version
    {
        get => Read(static part => part.CoreFileProperties.Version);
        set => SetString(value, static (part, text) => part.CoreFileProperties.Version = text);
    }

    private void SetString(string? value, Action<CoreFilePropertiesPart, string?> write)
    {
        string? validated = Validate(value);
        CoreFilePropertiesPart? part = CorePartForWrite(validated is not null);
        if (part is not null)
        {
            Write(part, property => write(property, validated));
        }
    }

    private void SetDate(DateTime? value, Action<CoreFilePropertiesPart, DateTime?> write)
    {
        CoreFilePropertiesPart? part = CorePartForWrite(value.HasValue);
        if (part is not null)
        {
            Write(part, property => write(property, value));
        }
    }

    private T Read<T>(Func<CoreFilePropertiesPart, T> read)
    {
        CoreFilePropertiesPart? part = _document.Package.CoreFilePropertiesPart;
        if (part is null)
        {
            return default!;
        }

        try
        {
            return read(part);
        }
        catch (Exception ex) when (IsMalformedCoreProperties(ex))
        {
            throw new InvalidOfficeFileException("The document's core properties are not readable.", ex);
        }
    }

    private static void Write(CoreFilePropertiesPart part, Action<CoreFilePropertiesPart> write)
    {
        try
        {
            write(part);
        }
        catch (Exception ex) when (IsMalformedCoreProperties(ex))
        {
            throw new InvalidOfficeFileException("The document's core properties are not readable.", ex);
        }
    }

    private CoreFilePropertiesPart? CorePartForWrite(bool create)
    {
        WordprocessingDocument package = _document.Package;
        CoreFilePropertiesPart? part = package.CoreFilePropertiesPart;
        if (part is null && create)
        {
            part = package.AddNewPart<CoreFilePropertiesPart>();
            using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
            byte[] root = Encoding.UTF8.GetBytes(EmptyCorePropertiesXml);
            stream.Write(root, 0, root.Length);
        }

        return part;
    }

    private static bool IsMalformedCoreProperties(Exception exception)
        => exception is XmlException or FormatException or InvalidDataException or OpenXmlPackageException;

    private static string? Validate(string? value)
    {
        if (value is null)
        {
            return null;
        }

        try
        {
            XmlConvert.VerifyXmlChars(value);
        }
        catch (XmlException ex)
        {
            throw new ArgumentException(
                "The value contains a character that Office Open XML cannot store.",
                nameof(value),
                ex);
        }

        return value;
    }
}
