using OfficeEngine.LibreOffice;

namespace OfficeEngine.Tests.Support;

/// <summary>Produces files with the external engine, for the tests that compare against it.</summary>
internal static class LibreOfficeTestSupport
{
    /// <summary>
    /// The filter that exports the first sheet of a workbook as CSV. The language identifier is
    /// forced to English (1033) so that the booleans come out as TRUE and FALSE, like in Excel.
    /// </summary>
    private const string CsvFilter = "csv:Text - txt - csv (StarCalc):44,34,76,1,,1033,false,true,true,false,false,0";

    /// <summary>Converts a workbook to CSV, with the values recalculated by the engine.</summary>
    /// <param name="workbook">The bytes of the workbook.</param>
    /// <returns>The lines of the CSV, one per row of the first sheet.</returns>
    internal static string[] ToCsv(byte[] workbook)
    {
        string root = Path.Combine(Path.GetTempPath(), "OfficeEngine.Tests.Csv", Guid.NewGuid().ToString("N"));
        string inputDirectory = Path.Combine(root, "in");
        string outputDirectory = Path.Combine(root, "out");
        string profileDirectory = Path.Combine(root, "profile");
        Directory.CreateDirectory(inputDirectory);
        Directory.CreateDirectory(outputDirectory);
        Directory.CreateDirectory(profileDirectory);

        try
        {
            string sourcePath = Path.Combine(inputDirectory, "book.xlsx");
            File.WriteAllBytes(sourcePath, workbook);

            LibreOfficeEngine engine = new();
            LibreOfficeRunResult run = engine.Run(
                LibreOfficeEngine.ConversionArguments(profileDirectory, outputDirectory, sourcePath, CsvFilter),
                TimeSpan.FromMinutes(2));
            if (run.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    $"The engine failed to export the CSV: {string.Join(" | ", run.Messages)}");
            }

            string csvPath = Path.Combine(outputDirectory, "book.csv");
            if (!File.Exists(csvPath))
            {
                throw new InvalidOperationException("The engine did not write the CSV.");
            }

            return File.ReadAllLines(csvPath);
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
