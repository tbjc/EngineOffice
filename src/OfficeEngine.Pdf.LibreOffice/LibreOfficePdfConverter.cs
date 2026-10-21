using System.Diagnostics;
using OfficeEngine.LibreOffice;
using OfficeEngine.Pdf;

namespace OfficeEngine.Pdf.LibreOffice;

/// <summary>
/// Converts documents to PDF with LibreOffice in headless mode.
/// </summary>
/// <remarks>
/// <para>
/// The engine runs as an isolated process: it gets its own temporary directory and its own user
/// profile, so it never touches the profile of a LibreOffice the user has open, and two conversions
/// do not collide.
/// </para>
/// <para>
/// The rendering depends on the version of LibreOffice, on the installed fonts and on the
/// configuration of the system. The result is a PDF that a reader can open, not a page-for-page
/// copy of what Microsoft Office would print.
/// </para>
/// </remarks>
public sealed class LibreOfficePdfConverter : IPdfConverter, IDisposable
{
    private readonly LibreOfficeEngine _engine;
    private readonly SemaphoreSlim _gate;

    /// <summary>Initializes a new instance of the <see cref="LibreOfficePdfConverter"/> class.</summary>
    /// <param name="options">The settings of the converter. The defaults are used when it is not given.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The maximum number of concurrent conversions is less than one.
    /// </exception>
    public LibreOfficePdfConverter(LibreOfficePdfConverterOptions? options = null)
    {
        LibreOfficePdfConverterOptions effective = options ?? new LibreOfficePdfConverterOptions();
        if (effective.MaximumConcurrentConversions < 1)
        {
            throw new ArgumentOutOfRangeException(
                nameof(options),
                effective.MaximumConcurrentConversions,
                "At least one conversion must be allowed to run at the same time.");
        }

        _engine = new LibreOfficeEngine(new LibreOfficeEngineOptions
        {
            ExecutablePath = effective.ExecutablePath,
            MinimumVersion = effective.MinimumVersion,
            VersionCheckTimeout = effective.VersionCheckTimeout,
        });
        _gate = new SemaphoreSlim(
            effective.MaximumConcurrentConversions,
            effective.MaximumConcurrentConversions);
    }

    /// <summary>Gets whether the LibreOffice executable was found.</summary>
    /// <remarks>The check looks for the file; it does not start the engine.</remarks>
    public bool IsAvailable => _engine.IsAvailable;

    /// <summary>Gets the path of the LibreOffice executable that was found, or <see langword="null"/>.</summary>
    public string? ExecutablePath => _engine.ExecutablePath;

    /// <summary>Gets the version of the LibreOffice that was found, or <see langword="null"/> when it cannot be read.</summary>
    public string? Version => _engine.Version;

    /// <inheritdoc />
    public PdfConversionResult Convert(Stream source, PdfSourceFormat format, PdfConversionOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(source);
        if (!Enum.IsDefined(format))
        {
            throw new ArgumentOutOfRangeException(nameof(format), format, "The source format is not defined.");
        }

        PdfConversionOptions effective = options ?? new PdfConversionOptions();
        effective.Validate();
        if (!source.CanRead)
        {
            throw new ArgumentException("The source stream must be readable.", nameof(source));
        }

        byte[] content = ReadSource(source, effective.MaximumInputBytes);
        if (content.Length == 0)
        {
            throw new ArgumentException("The source document is empty.", nameof(source));
        }

        _gate.Wait();
        try
        {
            return RunConversion(content, format, effective);
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>Releases the resources of the converter.</summary>
    public void Dispose()
    {
        _gate.Dispose();
    }

    private PdfConversionResult RunConversion(byte[] content, PdfSourceFormat format, PdfConversionOptions options)
    {
        string root = Path.Combine(Path.GetTempPath(), "OfficeEngine.Pdf", Guid.NewGuid().ToString("N"));
        string inputDirectory = Path.Combine(root, "in");
        string outputDirectory = Path.Combine(root, "out");
        string profileDirectory = Path.Combine(root, "profile");
        Directory.CreateDirectory(inputDirectory);
        Directory.CreateDirectory(outputDirectory);
        Directory.CreateDirectory(profileDirectory);

        Stopwatch watch = Stopwatch.StartNew();
        try
        {
            string extension = format == PdfSourceFormat.WordDocument ? ".docx" : ".xlsx";
            string filter = format == PdfSourceFormat.WordDocument ? "pdf:writer_pdf_Export" : "pdf:calc_pdf_Export";
            string sourcePath = Path.Combine(inputDirectory, "source" + extension);
            File.WriteAllBytes(sourcePath, content);

            LibreOfficeRunResult run;
            try
            {
                run = _engine.Run(
                    LibreOfficeEngine.ConversionArguments(profileDirectory, outputDirectory, sourcePath, filter),
                    options.Timeout);
            }
            catch (Exception ex) when (ex is LibreOfficeUnavailableException or TimeoutException)
            {
                throw new PdfConversionException(ex.Message, null, [], ex);
            }

            if (run.ExitCode != 0)
            {
                throw new PdfConversionException(
                    $"LibreOffice finished the conversion with exit code {run.ExitCode}.",
                    run.ExitCode,
                    run.Messages);
            }

            string pdfPath = Path.Combine(outputDirectory, "source.pdf");
            if (!File.Exists(pdfPath))
            {
                throw new PdfConversionException(
                    "LibreOffice finished without producing a PDF.",
                    run.ExitCode,
                    run.Messages);
            }

            var info = new FileInfo(pdfPath);
            if (info.Length == 0)
            {
                throw new PdfConversionException(
                    "LibreOffice produced an empty PDF.",
                    run.ExitCode,
                    run.Messages);
            }

            if (info.Length > options.MaximumOutputBytes)
            {
                throw new OfficeResourceLimitExceededException(
                    $"The produced PDF is larger than the allowed {options.MaximumOutputBytes} bytes.");
            }

            byte[] pdf = File.ReadAllBytes(pdfPath);
            if (!LooksLikePdf(pdf))
            {
                throw new PdfConversionException(
                    "The file produced by LibreOffice is not a PDF.",
                    run.ExitCode,
                    run.Messages);
            }

            watch.Stop();
            return new PdfConversionResult(pdf, run.Messages, watch.Elapsed);
        }
        finally
        {
            if (!options.KeepTemporaryFiles)
            {
                TryDelete(root);
            }
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
                throw new OfficeResourceLimitExceededException(
                    $"The source document is larger than the allowed {maximumBytes} bytes.");
            }

            buffer.Write(chunk, 0, read);
        }

        return buffer.ToArray();
    }

    private static bool LooksLikePdf(byte[] pdf)
    {
        ReadOnlySpan<byte> header = "%PDF-"u8;
        return pdf.Length >= header.Length && pdf.AsSpan(0, header.Length).SequenceEqual(header);
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
