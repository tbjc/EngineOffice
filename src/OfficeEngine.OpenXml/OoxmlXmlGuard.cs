using System.IO.Compression;
using System.Xml;
using OfficeEngine;

namespace OfficeEngine.OpenXml;

/// <summary>
/// Validates the XML parts of a package before the OOXML reader parses them.
/// </summary>
/// <remarks>
/// <para>
/// The Open XML SDK loads a part with a <em>recursive</em> reader: a document with a couple of
/// thousand nested elements exhausts the stack and terminates the process with an error that cannot
/// be caught. Measuring the depth beforehand turns that into an ordinary, typed failure.
/// </para>
/// <para>
/// The parts are measured with a streaming reader, which does not use the call stack, and the check
/// also proves that the part is well-formed XML and declares no document type definition. A part that
/// is not XML is rejected here, with a typed error, instead of surfacing later as an exception of
/// the XML library.
/// </para>
/// </remarks>
internal static class OoxmlXmlGuard
{
    /// <summary>Ensures that the XML parts of the package are safe to hand to the reader.</summary>
    /// <param name="packageStream">A seekable stream positioned anywhere; the position is restored.</param>
    /// <param name="limits">The limits to enforce.</param>
    /// <exception cref="InvalidOfficeFileException">
    /// The stream is not a readable ZIP container, or a part is not well-formed XML.
    /// </exception>
    /// <exception cref="OfficeResourceLimitExceededException">A configured limit was exceeded.</exception>
    internal static void EnsureSafe(Stream packageStream, OfficeEngineLimits limits)
    {
        ArgumentNullException.ThrowIfNull(packageStream);
        ArgumentNullException.ThrowIfNull(limits);
        limits.EnsureValid();

        long origin = packageStream.Position;
        packageStream.Position = 0;
        try
        {
            using ZipArchive archive = new(packageStream, ZipArchiveMode.Read, leaveOpen: true);
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                if (entry.Length == 0)
                {
                    continue;
                }

                using Stream content = entry.Open();
                byte[] head = ReadHead(content);
                if (!LooksLikeXml(head))
                {
                    continue;
                }

                using MemoryStream buffer = new();
                buffer.Write(head, 0, head.Length);
                content.CopyTo(buffer);
                buffer.Position = 0;
                Scan(buffer, entry.FullName, limits.MaximumXmlDepth);
            }
        }
        catch (InvalidDataException ex)
        {
            throw new InvalidOfficeFileException("The input is not a valid Office Open XML package (it is not a readable ZIP container).", ex);
        }
        finally
        {
            packageStream.Position = origin;
        }
    }

    private static void Scan(Stream content, string entryName, int maximumDepth)
    {
        XmlReaderSettings settings = new()
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            CloseInput = false,
        };

        try
        {
            using XmlReader reader = XmlReader.Create(content, settings);
            while (reader.Read())
            {
                if (reader.Depth > maximumDepth)
                {
                    throw new OfficeResourceLimitExceededException(
                        $"The part '{entryName}' nests XML deeper than the configured maximum of {maximumDepth} levels.");
                }
            }
        }
        catch (XmlException ex)
        {
            throw new InvalidOfficeFileException($"The part '{entryName}' is not valid XML.", ex);
        }
    }

    /// <summary>
    /// Decides whether an entry is XML by looking at its first bytes, so that a part renamed to
    /// hide its type is still measured.
    /// </summary>
    private static bool LooksLikeXml(byte[] head)
    {
        foreach (byte value in head)
        {
            if (value is 0xEF or 0xBB or 0xBF)
            {
                // Byte order mark: keep looking for the first real byte.
                continue;
            }

            if (value is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n')
            {
                continue;
            }

            return value == (byte)'<';
        }

        return false;
    }

    private static byte[] ReadHead(Stream content)
    {
        byte[] head = new byte[64];
        int read = content.Read(head, 0, head.Length);
        if (read == head.Length)
        {
            return head;
        }

        byte[] trimmed = new byte[read];
        Array.Copy(head, trimmed, read);
        return trimmed;
    }
}
