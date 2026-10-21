using System.IO.Compression;
using System.Text;

namespace OfficeEngine.Tests.Support;

/// <summary>
/// Helpers that inspect produced packages with an independent tool (System.IO.Compression)
/// instead of through the library under test.
/// </summary>
internal static class TestPackage
{
    /// <summary>Returns everything in the buffer as a seekable stream.</summary>
    internal static MemoryStream ToStream(byte[] bytes) => new(bytes, writable: false);

    /// <summary>Reads a specific ZIP entry as text, failing the test when it is missing.</summary>
    internal static string ReadEntryText(byte[] package, string entryName)
    {
        using MemoryStream stream = new(package, writable: false);
        using ZipArchive archive = new(stream, ZipArchiveMode.Read);
        ZipArchiveEntry entry = archive.GetEntry(entryName)
            ?? throw new Xunit.Sdk.XunitException($"The package does not contain an entry named '{entryName}'.");

        using Stream entryStream = entry.Open();
        using StreamReader reader = new(entryStream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    /// <summary>Determines whether a ZIP entry exists.</summary>
    internal static bool HasEntry(byte[] package, string entryName)
    {
        using MemoryStream stream = new(package, writable: false);
        using ZipArchive archive = new(stream, ZipArchiveMode.Read);
        return archive.GetEntry(entryName) is not null;
    }

    /// <summary>Lists the entry names of the package.</summary>
    internal static IReadOnlyList<string> EntryNames(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using ZipArchive archive = new(stream, ZipArchiveMode.Read);
        return archive.Entries.Select(static entry => entry.FullName).ToList();
    }

    /// <summary>Rebuilds a package with a single entry replaced by the given text, so tests can produce
    /// malformed content the SDK would refuse to author directly.</summary>
    internal static byte[] WithEntryText(byte[] package, string entryName, string newText)
    {
        using MemoryStream source = new(package, writable: false);
        using ZipArchive archive = new(source, ZipArchiveMode.Read);
        using MemoryStream output = new();
        using (ZipArchive copy = new(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                ZipArchiveEntry newEntry = copy.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using Stream target = newEntry.Open();
                if (string.Equals(entry.FullName, entryName, StringComparison.Ordinal))
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(newText);
                    target.Write(bytes, 0, bytes.Length);
                }
                else
                {
                    using Stream original = entry.Open();
                    original.CopyTo(target);
                }
            }
        }

        return output.ToArray();
    }

    /// <summary>Rebuilds a package without the given entry, so tests can produce a broken package.</summary>
    internal static byte[] WithoutEntry(byte[] package, string entryName)
    {
        using MemoryStream source = new(package);
        using ZipArchive archive = new(source, ZipArchiveMode.Read);
        using MemoryStream output = new();
        using (ZipArchive copy = new(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                if (string.Equals(entry.FullName, entryName, StringComparison.Ordinal))
                {
                    continue;
                }

                ZipArchiveEntry newEntry = copy.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using Stream target = newEntry.Open();
                using Stream original = entry.Open();
                original.CopyTo(target);
            }
        }

        return output.ToArray();
    }

    /// <summary>Builds a ZIP container with a single, highly compressible entry.</summary>
    internal static byte[] CreateCompressibleZip(string entryName, int uncompressedSize)
    {
        using MemoryStream stream = new();
        using (ZipArchive archive = new(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using Stream entryStream = entry.Open();
            entryStream.Write(new byte[uncompressedSize]);
        }

        return stream.ToArray();
    }

    /// <summary>Builds a ZIP container with the requested number of tiny entries.</summary>
    internal static byte[] CreateZipWithEntryCount(int entryCount)
    {
        using MemoryStream stream = new();
        using (ZipArchive archive = new(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            for (int i = 0; i < entryCount; i++)
            {
                ZipArchiveEntry entry = archive.CreateEntry($"part{i}.bin", CompressionLevel.Optimal);
                using Stream entryStream = entry.Open();
                entryStream.WriteByte((byte)i);
            }
        }

        return stream.ToArray();
    }
}
