using System.Diagnostics;
using OfficeEngine.LibreOffice;

namespace OfficeEngine.Legacy;

/// <summary>
/// Converts the legacy Word and Excel formats into their modern counterpart with an external engine.
/// </summary>
/// <remarks>
/// <para>
/// This is a **conversion**, not a reader: the legacy file is handed to the engine, which produces a
/// <em>new</em> document in the modern format. The result is the engine's interpretation of the
/// original, not a byte-level translation, and the legacy format is never edited in place.
/// </para>
/// <para>
/// To work with the content, convert first and then open the result with
/// <c>WordDocument.Open</c> or <c>ExcelWorkbook.Open</c>. Reading, editing and saving the legacy
/// format are <em>not</em> supported: saving a converted document writes the modern format.
/// </para>
/// <para>
/// What the conversion loses depends on the engine and on the file: macros, embedded objects, OLE
/// links, the exact layout and the metadata of the original are not preserved. The conversion is a
/// best effort of the engine, not a guarantee.
/// </para>
/// </remarks>
public sealed class LegacyConverter
{
    private readonly LegacyConverterOptions _options;
    private readonly LibreOfficeEngine _engine;

    /// <summary>Initializes a new instance of the <see cref="LegacyConverter"/> class.</summary>
    /// <param name="options">The settings of the converter. The defaults are used when it is not given.</param>
    public LegacyConverter(LegacyConverterOptions? options = null)
    {
        _options = options ?? new LegacyConverterOptions();
        _engine = new LibreOfficeEngine(new LibreOfficeEngineOptions
        {
            ExecutablePath = _options.ExecutablePath,
            MinimumVersion = _options.MinimumVersion,
            VersionCheckTimeout = _options.VersionCheckTimeout,
        });
    }

    /// <summary>Gets whether the LibreOffice executable was found.</summary>
    /// <remarks>The check looks for the file; it does not start the engine.</remarks>
    public bool IsAvailable => _engine.IsAvailable;

    /// <summary>Gets the path of the LibreOffice executable that was found, or <see langword="null"/>.</summary>
    public string? ExecutablePath => _engine.ExecutablePath;

    /// <summary>Gets the version of the LibreOffice that was found, or <see langword="null"/>.</summary>
    public string? Version => _engine.Version;

    /// <summary>Converts a legacy document into its modern counterpart.</summary>
    /// <param name="source">The stream with the legacy document. It is read from its current position.</param>
    /// <param name="format">The format of the legacy document.</param>
    /// <returns>The converted document and the diagnostics of the engine.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// The stream is empty, is not readable or is larger than the allowed size.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is not a defined value.</exception>
    /// <exception cref="LegacyConversionException">
    /// The engine is not available, the conversion failed or the result is not a modern document.
    /// </exception>
    public LegacyConversionResult Convert(Stream source, LegacyDocumentFormat format)
    {
        ArgumentNullException.ThrowIfNull(source);
        if (!Enum.IsDefined(format))
        {
            throw new ArgumentOutOfRangeException(nameof(format), format, "The legacy format is not defined.");
        }

        if (!source.CanRead)
        {
            throw new ArgumentException("The source stream must be readable.", nameof(source));
        }

        byte[] content = ReadSource(source, _options.MaximumInputBytes);
        if (content.Length == 0)
        {
            throw new ArgumentException("The source document is empty.", nameof(source));
        }

        string root = Path.Combine(Path.GetTempPath(), "OfficeEngine.Legacy", Guid.NewGuid().ToString("N"));
        string inputDirectory = Path.Combine(root, "in");
        string outputDirectory = Path.Combine(root, "out");
        string profileDirectory = Path.Combine(root, "profile");
        Directory.CreateDirectory(inputDirectory);
        Directory.CreateDirectory(outputDirectory);
        Directory.CreateDirectory(profileDirectory);

        Stopwatch watch = Stopwatch.StartNew();
        try
        {
            (string extension, string filter, string targetExtension) = format switch
            {
                LegacyDocumentFormat.Word97To2003 => (".doc", "docx:MS Word 2007 XML", ".docx"),
                LegacyDocumentFormat.Excel97To2003 => (".xls", "xlsx:Calc MS Excel 2007 XML", ".xlsx"),
                _ => throw new ArgumentOutOfRangeException(nameof(format), format, "The legacy format is not defined."),
            };

            string sourcePath = Path.Combine(inputDirectory, "source" + extension);
            File.WriteAllBytes(sourcePath, content);

            LibreOfficeRunResult run;
            try
            {
                run = _engine.Run(
                    LibreOfficeEngine.ConversionArguments(profileDirectory, outputDirectory, sourcePath, filter),
                    _options.Timeout);
            }
            catch (Exception ex) when (ex is LibreOfficeUnavailableException or TimeoutException)
            {
                throw new LegacyConversionException(ex.Message, innerException: ex);
            }

            if (run.ExitCode != 0)
            {
                throw new LegacyConversionException(
                    $"The engine finished the conversion with exit code {run.ExitCode}.",
                    run.ExitCode,
                    run.Messages);
            }

            string targetPath = Path.Combine(outputDirectory, "source" + targetExtension);
            if (!File.Exists(targetPath))
            {
                throw new LegacyConversionException(
                    "The engine finished without producing a converted document.",
                    run.ExitCode,
                    run.Messages);
            }

            byte[] converted = File.ReadAllBytes(targetPath);
            if (!LooksLikePackage(converted))
            {
                throw new LegacyConversionException(
                    "The file produced by the engine is not an Office Open XML package.",
                    run.ExitCode,
                    run.Messages);
            }

            watch.Stop();
            return new LegacyConversionResult(converted, format, run.Messages, watch.Elapsed);
        }
        finally
        {
            TryDelete(root);
        }
    }

    private static byte[] ReadSource(Stream source, long maximumBytes)
    {
        using MemoryStream buffer = new();
        byte[] chunk = new byte[81920];
        int read;
        while ((read = source.Read(chunk, 0, chunk.Length)) > 0)
        {
            if (buffer.Length + read > maximumBytes)
            {
                throw new LegacyConversionException(
                    $"The legacy document is larger than the allowed {maximumBytes} bytes.");
            }

            buffer.Write(chunk, 0, read);
        }

        return buffer.ToArray();
    }

    private static bool LooksLikePackage(byte[] content)
    {
        ReadOnlySpan<byte> signature = "PK\x03\x04"u8;
        return content.Length >= signature.Length && content.AsSpan(0, signature.Length).SequenceEqual(signature);
    }

    private static void TryDelete(string directory)
    {
        try
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, recursive: true);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // The temporary files stay; the conversion itself already succeeded or failed.
        }
    }
}
