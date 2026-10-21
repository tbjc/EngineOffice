using System.Buffers;
using OfficeEngine;

namespace OfficeEngine.OpenXml;

/// <summary>
/// Reads untrusted input into a bounded, seekable in-memory buffer.
/// </summary>
/// <remarks>
/// OfficeEngine keeps the package contents in memory for the lifetime of an open document.
/// This trades memory for a simple, file-lock-free lifecycle. The maximum size is enforced
/// here so that a hostile or accidental large input cannot exhaust memory.
/// </remarks>
internal static class OoxmlLoader
{
    private const int BufferSize = 81920;

    /// <summary>Loads a file into a bounded in-memory buffer.</summary>
    /// <param name="path">Path of the file to read.</param>
    /// <param name="limits">Limits to enforce.</param>
    /// <returns>A seekable buffer positioned at the beginning.</returns>
    internal static MemoryStream LoadFromFile(string path, OfficeEngineLimits limits)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(limits);
        limits.EnsureValid();

        using FileStream file = new(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (file.Length > limits.MaxFileSizeBytes)
        {
            throw new OfficeResourceLimitExceededException(
                $"The file is {file.Length} bytes, which exceeds the configured maximum of {limits.MaxFileSizeBytes} bytes.");
        }

        return ReadBounded(file, limits);
    }

    /// <summary>Loads a stream into a bounded in-memory buffer.</summary>
    /// <param name="source">The stream to read.</param>
    /// <param name="leaveOpen">
    /// When <see langword="false"/> (the default), <paramref name="source"/> is disposed after it has been read.
    /// </param>
    /// <param name="limits">Limits to enforce.</param>
    /// <returns>A seekable buffer positioned at the beginning.</returns>
    internal static MemoryStream LoadFromStream(Stream source, bool leaveOpen, OfficeEngineLimits limits)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(limits);
        limits.EnsureValid();

        MemoryStream buffer = ReadBounded(source, limits);
        if (!leaveOpen)
        {
            source.Dispose();
        }

        return buffer;
    }

    private static MemoryStream ReadBounded(Stream source, OfficeEngineLimits limits)
    {
        if (source.CanSeek && source.Length - source.Position > limits.MaxFileSizeBytes)
        {
            throw new OfficeResourceLimitExceededException(
                $"The input is larger than the configured maximum of {limits.MaxFileSizeBytes} bytes.");
        }

        MemoryStream buffer = new();
        byte[] rented = ArrayPool<byte>.Shared.Rent(BufferSize);
        try
        {
            long total = 0;
            int read;
            while ((read = source.Read(rented, 0, rented.Length)) > 0)
            {
                total += read;
                if (total > limits.MaxFileSizeBytes)
                {
                    throw new OfficeResourceLimitExceededException(
                        $"The input is larger than the configured maximum of {limits.MaxFileSizeBytes} bytes.");
                }

                buffer.Write(rented, 0, read);
            }
        }
        catch
        {
            buffer.Dispose();
            throw;
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(rented);
        }

        buffer.Position = 0;
        return buffer;
    }
}
