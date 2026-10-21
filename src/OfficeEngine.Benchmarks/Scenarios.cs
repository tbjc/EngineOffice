using System.Globalization;
using System.Text;
using OfficeEngine.Excel;
using OfficeEngine.Word;

namespace OfficeEngine.Benchmarks;

/// <summary>The scenarios that are measured, and the baseline they are compared against.</summary>
internal static class Scenarios
{
    internal const int WordParagraphs = 5000;
    internal const int ExcelRows = 5000;

    internal static List<BenchmarkResult> Run()
    {
        List<BenchmarkResult> results = [];

        byte[] wordPackage = [];
        results.Add(Benchmark.Measure("word.create", () =>
        {
            using WordDocument document = WordDocument.Create();
            for (int index = 0; index < WordParagraphs; index++)
            {
                document.AddParagraph($"Párrafo {index} con un texto de relleno para la medición.");
            }

            using MemoryStream stream = new();
            document.Save(stream);
            wordPackage = stream.ToArray();
        }, () => wordPackage.Length));

        results.Add(Benchmark.Measure("word.open", () =>
        {
            using MemoryStream stream = new(wordPackage);
            using WordDocument document = WordDocument.Open(stream);
            _ = document.Paragraphs.Count;
        }));

        results.Add(Benchmark.Measure("word.roundtrip", () =>
        {
            using MemoryStream source = new(wordPackage);
            using WordDocument document = WordDocument.Open(source);
            using MemoryStream target = new();
            document.Save(target);
        }));

        byte[] excelPackage = [];
        results.Add(Benchmark.Measure("excel.create", () =>
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Create();
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 1; row <= ExcelRows; row++)
            {
                sheet[row, 1].Value = $"Fila {row}";
                sheet[row, 2].Value = row * 1.5d;
            }

            using MemoryStream stream = new();
            workbook.Save(stream);
            excelPackage = stream.ToArray();
        }, () => excelPackage.Length));

        results.Add(Benchmark.Measure("excel.open", () =>
        {
            using MemoryStream stream = new(excelPackage);
            using ExcelWorkbook workbook = ExcelWorkbook.Open(stream);
            _ = workbook.Worksheets.Count;
        }));

        results.Add(Benchmark.Measure("excel.roundtrip", () =>
        {
            using MemoryStream source = new(excelPackage);
            using ExcelWorkbook workbook = ExcelWorkbook.Open(source);
            using MemoryStream target = new();
            workbook.Save(target);
        }));

        return results;
    }

    internal static string ToBaseline(IEnumerable<BenchmarkResult> results)
    {
        StringBuilder builder = new();
        foreach (BenchmarkResult result in results)
        {
            builder.Append(CultureInfo.InvariantCulture, $"{result.Name}|{result.Elapsed.TotalMilliseconds:0.0}|{result.AllocatedBytes}|{result.PackageBytes}");
            builder.AppendLine();
        }

        return builder.ToString();
    }

    internal static Dictionary<string, (double Milliseconds, long Allocated, long Package)> ParseBaseline(string text)
    {
        Dictionary<string, (double, long, long)> baseline = [];
        foreach (string line in text.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            string[] parts = line.Split('|');
            if (parts.Length != 4)
            {
                continue;
            }

            if (double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double milliseconds)
                && long.TryParse(parts[2], out long allocated)
                && long.TryParse(parts[3], out long package))
            {
                baseline[parts[0]] = (milliseconds, allocated, package);
            }
        }

        return baseline;
    }
}
