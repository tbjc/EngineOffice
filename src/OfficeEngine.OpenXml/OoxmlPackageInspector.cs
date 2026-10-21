using System.IO.Compression;
using OfficeEngine;

namespace OfficeEngine.OpenXml;

/// <summary>
/// Performs cheap, non-decompressing checks on an Open Packaging Conventions container
/// before it is handed to the OOXML reader.
/// </summary>
/// <remarks>
/// The checks are based purely on the ZIP central directory, so they never decompress
/// entry data and therefore cannot themselves trigger a decompression bomb.
/// </remarks>
internal static class OoxmlPackageInspector
{
    /// <summary>Ensures the package does not exceed the configured limits.</summary>
    /// <param name="packageStream">A seekable stream positioned anywhere; the position is restored.</param>
    /// <param name="limits">The limits to enforce.</param>
    /// <exception cref="InvalidOfficeFileException">The stream is not a readable ZIP container.</exception>
    /// <exception cref="OfficeResourceLimitExceededException">A configured limit was exceeded.</exception>
    internal static void EnsureWithinLimits(Stream packageStream, OfficeEngineLimits limits)
    {
        ArgumentNullException.ThrowIfNull(packageStream);
        ArgumentNullException.ThrowIfNull(limits);
        limits.EnsureValid();

        long origin = packageStream.Position;
        packageStream.Position = 0;
        try
        {
            using ZipArchive archive = new(packageStream, ZipArchiveMode.Read, leaveOpen: true);

            if (archive.Entries.Count > limits.MaxEntryCount)
            {
                throw new OfficeResourceLimitExceededException(
                    $"The package contains {archive.Entries.Count} entries, which exceeds the configured maximum of {limits.MaxEntryCount}.");
            }

            long totalUncompressed = 0;
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                long uncompressed = entry.Length;
                totalUncompressed += uncompressed;
                if (totalUncompressed > limits.MaxTotalUncompressedSizeBytes)
                {
                    throw new OfficeResourceLimitExceededException(
                        $"The package expands to more than the configured maximum of {limits.MaxTotalUncompressedSizeBytes} bytes.");
                }

                long compressed = entry.CompressedLength;
                double ratio = compressed > 0 ? uncompressed / (double)compressed : uncompressed > 0 ? double.PositiveInfinity : 1.0;
                if (ratio > limits.MaxEntryCompressionRatio)
                {
                    throw new OfficeResourceLimitExceededException(
                        $"Entry '{entry.FullName}' has a compression ratio of {ratio:0.##}, which exceeds the configured maximum of {limits.MaxEntryCompressionRatio}.");
                }
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
}
