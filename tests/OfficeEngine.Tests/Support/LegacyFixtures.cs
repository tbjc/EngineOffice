using OfficeEngine.LibreOffice;
using OfficeEngine.Legacy;

namespace OfficeEngine.Tests.Support;

/// <summary>
/// Produces the legacy files that the conversion tests need, with the same engine that converts
/// them back.
/// </summary>
internal static class LegacyFixtures
{
    internal static byte[] ToLegacyWord(byte[] docx)
        => Convert(docx, "docx", "doc:MS Word 97");

    internal static byte[] ToLegacyExcel(byte[] xlsx)
        => Convert(xlsx, "xlsx", "xls:MS Excel 97");

    private static byte[] Convert(byte[] content, string sourceExtension, string filter)
    {
        string root = Path.Combine(Path.GetTempPath(), "OfficeEngine.Tests.Legacy", Guid.NewGuid().ToString("N"));
        string inputDirectory = Path.Combine(root, "in");
        string outputDirectory = Path.Combine(root, "out");
        string profileDirectory = Path.Combine(root, "profile");
        Directory.CreateDirectory(inputDirectory);
        Directory.CreateDirectory(outputDirectory);
        Directory.CreateDirectory(profileDirectory);

        try
        {
            string sourcePath = Path.Combine(inputDirectory, "source." + sourceExtension);
            File.WriteAllBytes(sourcePath, content);

            LibreOfficeEngine engine = new();
            LibreOfficeRunResult run = engine.Run(
                LibreOfficeEngine.ConversionArguments(profileDirectory, outputDirectory, sourcePath, filter),
                TimeSpan.FromMinutes(2));
            if (run.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    $"The engine failed to write the legacy fixture: {string.Join(" | ", run.Messages)}");
            }

            string targetPath = Path.Combine(outputDirectory, "source." + (sourceExtension == "docx" ? "doc" : "xls"));
            if (!File.Exists(targetPath))
            {
                throw new InvalidOperationException("The engine did not write the legacy fixture.");
            }

            return File.ReadAllBytes(targetPath);
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // The temporary files stay; the test does not depend on them.
            }
        }
    }
}
