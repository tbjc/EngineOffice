namespace OfficeEngine.LibreOffice;

/// <summary>The settings of the LibreOffice process runner.</summary>
internal sealed class LibreOfficeEngineOptions
{
    /// <summary>
    /// The path of the LibreOffice executable. When it is not given, the usual installation
    /// directories and the <c>PATH</c> are searched.
    /// </summary>
    internal string? ExecutablePath { get; set; }

    /// <summary>
    /// The lowest version of LibreOffice that the runner accepts, as <c>major.minor</c>. The default
    /// is <c>7.0</c>. A build whose version cannot be read is accepted.
    /// </summary>
    internal string MinimumVersion { get; set; } = "7.0";

    /// <summary>How long the check of the version may take. The default is thirty seconds.</summary>
    internal TimeSpan VersionCheckTimeout { get; set; } = TimeSpan.FromSeconds(30);
}
