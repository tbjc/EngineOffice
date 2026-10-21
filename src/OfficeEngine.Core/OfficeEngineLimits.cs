namespace OfficeEngine;

/// <summary>
/// Resource limits applied while reading untrusted Office Open XML files.
/// </summary>
/// <remarks>
/// Office files are untrusted input. These limits are enforced before any content is
/// parsed and are intended to bound memory use and to detect maliciously crafted
/// packages (ZIP bombs, huge embedded parts).
/// <para>
/// Only limits that are actually enforced are exposed here. Each limit is verified by
/// a test in <c>OfficeEngine.Tests</c>. Limits that are not yet enforced (for example
/// wall-clock timeouts) are intentionally not part of this type.
/// </para>
/// </remarks>
public sealed class OfficeEngineLimits
{
    /// <summary>
    /// Gets a shared instance with conservative, server-friendly defaults.
    /// </summary>
    public static OfficeEngineLimits Default { get; } = new();

    /// <summary>
    /// Gets the maximum size, in bytes, of an input file or stream that will be read.
    /// Defaults to 256 MiB.
    /// </summary>
    public long MaxFileSizeBytes { get; init; } = 256L * 1024 * 1024;

    /// <summary>
    /// Gets the maximum number of entries (parts) allowed inside the package.
    /// Defaults to 4096.
    /// </summary>
    public int MaxEntryCount { get; init; } = 4096;

    /// <summary>
    /// Gets the maximum total uncompressed size, in bytes, of all entries in the package.
    /// Defaults to 512 MiB.
    /// </summary>
    public long MaxTotalUncompressedSizeBytes { get; init; } = 512L * 1024 * 1024;

    /// <summary>
    /// Gets the maximum allowed ratio between the uncompressed and compressed size of a
    /// single entry. Used to detect ZIP bombs. Defaults to 100.
    /// </summary>
    public double MaxEntryCompressionRatio { get; init; } = 100.0;

    /// <summary>
    /// Gets the maximum allowed nesting depth of the XML parts of the package.
    /// Defaults to 256.
    /// </summary>
    /// <remarks>
    /// The Open XML SDK loads a part with a recursive reader, so a deeply nested document
    /// exhausts the stack and terminates the process with an error that cannot be caught.
    /// The parts are measured before they are parsed, and a document deeper than this limit
    /// is rejected. Real Word and Excel documents are an order of magnitude below it.
    /// </remarks>
    public int MaximumXmlDepth { get; init; } = 256;

    /// <summary>
    /// Validates that the configured limits are usable.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">A limit has an invalid value.</exception>
    public void EnsureValid()
    {
        if (MaxFileSizeBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxFileSizeBytes), MaxFileSizeBytes, "The maximum file size must be greater than zero.");
        }

        if (MaxEntryCount <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxEntryCount), MaxEntryCount, "The maximum entry count must be greater than zero.");
        }

        if (MaxTotalUncompressedSizeBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxTotalUncompressedSizeBytes), MaxTotalUncompressedSizeBytes, "The maximum total uncompressed size must be greater than zero.");
        }

        if (MaxEntryCompressionRatio <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxEntryCompressionRatio), MaxEntryCompressionRatio, "The maximum compression ratio must be greater than zero.");
        }

        if (MaximumXmlDepth <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaximumXmlDepth), MaximumXmlDepth, "The maximum XML depth must be greater than zero.");
        }
    }
}
