using System.Diagnostics;
using System.Globalization;

namespace OfficeEngine.LibreOffice;

/// <summary>
/// Runs LibreOffice in headless mode as an isolated process.
/// </summary>
/// <remarks>
/// <para>
/// The engine is found in the usual installation directories or in the <c>PATH</c>, and on Windows the
/// console companion (<c>soffice.com</c>) is preferred, because the graphical executable detaches:
/// it does not wait and it does not report the exit code.
/// </para>
/// <para>
/// The process runs with its own user profile and its own temporary directory, so it never touches
/// the profile of a LibreOffice the user has open and two runs do not collide.
/// </para>
/// </remarks>
internal sealed class LibreOfficeEngine
{
    private const string WindowsExecutable = "soffice.exe";
    private const string WindowsConsoleExecutable = "soffice.com";
    private const string UnixExecutable = "soffice";

    private readonly LibreOfficeEngineOptions _options;
    private EngineInfo? _engine;

    internal LibreOfficeEngine(LibreOfficeEngineOptions? options = null)
        => _options = options ?? new LibreOfficeEngineOptions();

    /// <summary>Gets whether the LibreOffice executable was found. The check does not start it.</summary>
    internal bool IsAvailable => ResolveExecutable() is not null;

    /// <summary>Gets the path of the LibreOffice executable that was found, or <see langword="null"/>.</summary>
    internal string? ExecutablePath => ResolveExecutable();

    /// <summary>Gets the version of the LibreOffice that was found, or <see langword="null"/>.</summary>
    internal string? Version => Detect().Version;

    /// <summary>
    /// Verifies that the engine is usable and returns the path of its console executable.
    /// </summary>
    /// <exception cref="LibreOfficeUnavailableException">The engine is missing or too old.</exception>
    internal string EnsureAvailable()
    {
        EngineInfo engine = Detect();
        if (engine.ConsolePath is null)
        {
            throw new LibreOfficeUnavailableException(
                "LibreOffice was not found. Install it or point the adapter at the executable.");
        }

        if (engine.Version is not null && !MeetsMinimumVersion(engine.Version))
        {
            throw new LibreOfficeUnavailableException(
                $"LibreOffice {engine.Version} is older than the required version {_options.MinimumVersion}.");
        }

        return engine.ConsolePath;
    }

    /// <summary>Runs the engine with the given arguments and waits for it to finish.</summary>
    /// <param name="arguments">The arguments of the process.</param>
    /// <param name="timeout">How long the process may take.</param>
    /// <param name="workingDirectory">The working directory of the process.</param>
    /// <returns>What the process wrote and how it finished.</returns>
    /// <exception cref="TimeoutException">The process did not finish in time and was killed.</exception>
    internal LibreOfficeRunResult Run(IEnumerable<string> arguments, TimeSpan timeout, string? workingDirectory = null)
    {
        string executable = EnsureAvailable();
        ProcessStartInfo startInfo = new(executable)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = workingDirectory ?? string.Empty,
        };

        foreach (string argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using Process process = Process.Start(startInfo)!;
        Task<string> standardOutput = process.StandardOutput.ReadToEndAsync();
        Task<string> standardError = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit((int)timeout.TotalMilliseconds))
        {
            Kill(process);
            throw new TimeoutException($"LibreOffice did not finish within {timeout}.");
        }

        return new LibreOfficeRunResult(
            process.ExitCode,
            standardOutput.GetAwaiter().GetResult(),
            standardError.GetAwaiter().GetResult());
    }

    /// <summary>
    /// Builds the arguments that run the engine in headless mode with its own profile.
    /// </summary>
    /// <param name="profileDirectory">The directory used as the user profile of the engine.</param>
    /// <param name="outputDirectory">The directory where the engine writes its result.</param>
    /// <param name="sourcePath">The file the engine converts.</param>
    /// <param name="filter">The export filter, for example <c>pdf:writer_pdf_Export</c>.</param>
    /// <returns>The arguments of the process.</returns>
    internal static IEnumerable<string> ConversionArguments(
        string profileDirectory,
        string outputDirectory,
        string sourcePath,
        string filter)
    {
        yield return "--headless";
        yield return "--norestore";
        yield return "--nolockcheck";
        yield return "--nodefault";
        yield return "--nofirststartwizard";
        yield return "-env:UserInstallation=" + new Uri(profileDirectory).AbsoluteUri;
        yield return "--convert-to";
        yield return filter;
        yield return "--outdir";
        yield return outputDirectory;
        yield return sourcePath;
    }

    private EngineInfo Detect()
    {
        if (_engine is not null)
        {
            return _engine;
        }

        string? path = ResolveExecutable();
        _engine = new EngineInfo(path, path is null ? null : ReadVersion(path));
        return _engine;
    }

    private string? ResolveExecutable()
    {
        if (_options.ExecutablePath is string configured)
        {
            return File.Exists(configured) ? PreferConsoleExecutable(configured) : null;
        }

        foreach (string candidate in DefaultLocations())
        {
            if (File.Exists(candidate))
            {
                return PreferConsoleExecutable(candidate);
            }
        }

        string? pathVariable = Environment.GetEnvironmentVariable("PATH");
        if (pathVariable is null)
        {
            return null;
        }

        string fileName = OperatingSystem.IsWindows() ? WindowsExecutable : UnixExecutable;
        foreach (string directory in pathVariable.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            string candidate = Path.Combine(directory.Trim(), fileName);
            if (File.Exists(candidate))
            {
                return PreferConsoleExecutable(candidate);
            }
        }

        return null;
    }

    private static IEnumerable<string> DefaultLocations()
    {
        if (!OperatingSystem.IsWindows())
        {
            yield return "/usr/bin/soffice";
            yield return "/usr/local/bin/soffice";
            yield return "/opt/libreoffice/program/soffice";
            yield break;
        }

        foreach (string programFiles in new[]
        {
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
        })
        {
            if (string.IsNullOrEmpty(programFiles))
            {
                continue;
            }

            yield return Path.Combine(programFiles, "LibreOffice", "program", WindowsExecutable);
        }
    }

    /// <summary>
    /// On Windows the graphical executable detaches and does not report the exit code, so the
    /// console companion is preferred when it exists.
    /// </summary>
    private static string PreferConsoleExecutable(string path)
    {
        if (!OperatingSystem.IsWindows())
        {
            return path;
        }

        string directory = Path.GetDirectoryName(path) ?? string.Empty;
        string console = Path.Combine(directory, WindowsConsoleExecutable);
        return File.Exists(console) ? console : path;
    }

    private string? ReadVersion(string executable)
    {
        try
        {
            ProcessStartInfo startInfo = new(executable)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            startInfo.ArgumentList.Add("--version");

            using Process process = Process.Start(startInfo)!;
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit((int)_options.VersionCheckTimeout.TotalMilliseconds);
            foreach (string line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                const string prefix = "LibreOffice ";
                if (line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    int end = line.IndexOf(' ', prefix.Length);
                    return end < 0 ? line[prefix.Length..] : line[prefix.Length..end];
                }
            }
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException or IOException)
        {
            // The version cannot be read; the run still reports it as unknown.
        }

        return null;
    }

    private bool MeetsMinimumVersion(string version)
    {
        if (!TryParseVersion(version, out int major, out int minor)
            || !TryParseVersion(_options.MinimumVersion, out int requiredMajor, out int requiredMinor))
        {
            return true;
        }

        return major > requiredMajor || (major == requiredMajor && minor >= requiredMinor);
    }

    private static bool TryParseVersion(string value, out int major, out int minor)
    {
        major = 0;
        minor = 0;
        string[] parts = value.Split('.', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0
            || !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out major))
        {
            return false;
        }

        minor = parts.Length > 1
            && int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out int parsed)
                ? parsed
                : 0;
        return true;
    }

    private static void Kill(Process process)
    {
        try
        {
            process.Kill(entireProcessTree: true);
        }
        catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            // The process already finished; nothing else to do.
        }
    }

    private sealed class EngineInfo
    {
        internal EngineInfo(string? consolePath, string? version)
        {
            ConsolePath = consolePath;
            Version = version;
        }

        internal string? ConsolePath { get; }

        internal string? Version { get; }
    }
}
