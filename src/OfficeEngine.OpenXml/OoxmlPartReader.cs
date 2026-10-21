using System.IO;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeEngine.OpenXml;

/// <summary>
/// Materialises the DOM of a part, translating the failures of the reader into the typed error of
/// the library.
/// </summary>
/// <remarks>
/// The Open XML SDK reports a part that cannot be parsed with exceptions of its own (or of the XML
/// library). OfficeEngine promises typed errors, so those are translated here and the original
/// exception is kept as the inner one.
/// </remarks>
internal static class OoxmlPartReader
{
    /// <summary>Reads the DOM of a part, or throws <see cref="InvalidOfficeFileException"/>.</summary>
    /// <typeparam name="T">The type of the element that is read.</typeparam>
    /// <param name="partName">The name of the part, used in the message.</param>
    /// <param name="read">The accessor that materialises the DOM.</param>
    /// <returns>The element of the part.</returns>
    /// <exception cref="InvalidOfficeFileException">The part cannot be read as OOXML.</exception>
    internal static T Read<T>(string partName, Func<T?> read)
        where T : class
        => ReadOrDefault(partName, read)
            ?? throw new InvalidOfficeFileException($"The part '{partName}' has no root element.");

    /// <summary>Reads the DOM of a part that is allowed to have no root element.</summary>
    /// <typeparam name="T">The type of the element that is read.</typeparam>
    /// <param name="partName">The name of the part, used in the message.</param>
    /// <param name="read">The accessor that materialises the DOM.</param>
    /// <returns>The element of the part, or <see langword="null"/> when it has none.</returns>
    /// <exception cref="InvalidOfficeFileException">The part cannot be read as OOXML.</exception>
    internal static T? ReadOrDefault<T>(string partName, Func<T?> read)
        where T : class
    {
        try
        {
            return read();
        }
        catch (Exception ex) when (ex is InvalidDataException or XmlException or OpenXmlPackageException or FileFormatException
                or ArgumentOutOfRangeException)
        {
            throw new InvalidOfficeFileException($"The part '{partName}' cannot be read as OOXML.", ex);
        }
    }
}
