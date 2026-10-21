namespace OfficeEngine.Legacy;

/// <summary>The settings of the legacy converter.</summary>
public sealed class LegacyConverterOptions
{
    /// <summary>
    /// The path of the LibreOffice executable. When it is not given, the usual installation
    /// directories and the <c>PATH</c> are searched.
    /// </summary>
    public string? ExecutablePath { get; set; }

    /// <summary>
    /// The lowest version of LibreOffice that the converter accepts, as <c>major.minor</c>. The
    /// default is <c>7.0</c>. A build whose version cannot be read is accepted.
    /// </summary>
    public string MinimumVersion { get; set; } = "7.0";

    /// <summary>How long the check of the version may take. The default is thirty seconds.</summary>
    public TimeSpan VersionCheckTimeout { get; set; } = TimeSpan.FromSeconds(30);

    /// <summary>How long the conversion may take. The default is two minutes.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>
    /// The maximum size of the legacy document, in bytes. The default is 64 MiB.
    /// </summary>
    public long MaximumInputBytes { get; init; } = 64L * 1024 * 1024;
}
