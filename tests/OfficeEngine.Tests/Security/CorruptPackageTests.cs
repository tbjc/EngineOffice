using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Security;

/// <summary>
/// A corpus of mutated packages: every one of them must fail with a typed error or open, but never
/// with an exception of the SDK or of the XML library, and never with a crash.
/// </summary>
public sealed class CorruptPackageTests
{
    [Fact]
    public void Word_WithMutatedPackages_FailsWithATypedErrorOrOpens()
    {
        foreach (byte[] mutated in Mutations(CreateWordDocument(), "word/document.xml", "word/_rels/document.xml.rels"))
        {
            Exception? error = Record(() =>
            {
                using WordDocument document = WordDocument.Open(TestPackage.ToStream(mutated));
                _ = document.Paragraphs.Count;
            });

            if (error is not null)
            {
                Assert.True(
                    error is OfficeEngineException,
                    $"{error.GetType().FullName}: {error.Message}{Environment.NewLine}{error.StackTrace}");
            }
        }
    }

    [Fact]
    public void Excel_WithMutatedPackages_FailsWithATypedErrorOrOpens()
    {
        foreach (byte[] mutated in Mutations(CreateExcelWorkbook(), "xl/worksheets/sheet1.xml", "xl/_rels/workbook.xml.rels"))
        {
            Exception? error = Record(() =>
            {
                using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(mutated));
                _ = workbook.Worksheets.Count;
            });

            if (error is not null)
            {
                Assert.True(
                    error is OfficeEngineException,
                    $"{error.GetType().FullName}: {error.Message}{Environment.NewLine}{error.StackTrace}");
            }
        }
    }

    [Fact]
    public void TheCorpus_IsBigEnoughToCoverTheFailureModes()
    {
        List<byte[]> mutations = [.. Mutations(CreateWordDocument(), "word/document.xml", "word/_rels/document.xml.rels")];

        Assert.True(mutations.Count >= 20, $"The corpus only has {mutations.Count} mutations.");
    }

    private static IEnumerable<byte[]> Mutations(byte[] package, string part, string relationships)
    {
        // Truncations at ten points of the file.
        for (int fraction = 1; fraction <= 9; fraction++)
        {
            int length = package.Length * fraction / 10;
            if (length > 0 && length < package.Length)
            {
                yield return package[..length];
            }
        }

        // A flipped byte at ten points of the file.
        int step = Math.Max(1, package.Length / 10);
        for (int index = 0; index < package.Length; index += step)
        {
            byte[] copy = (byte[])package.Clone();
            copy[index] ^= 0xFF;
            yield return copy;
        }

        // Missing entries.
        yield return TestPackage.WithoutEntry(package, part);
        yield return TestPackage.WithoutEntry(package, relationships);
        yield return TestPackage.WithoutEntry(package, "[Content_Types].xml");

        // Entries replaced by content that is not what they claim to be.
        yield return TestPackage.WithEntryText(package, part, "esto no es xml");
        yield return TestPackage.WithEntryText(package, relationships, "<no-es-xml/>");
        yield return TestPackage.WithEntryText(package, "[Content_Types].xml", "<no-es-xml/>");
    }

    private static Exception? Record(Action action)
    {
        try
        {
            action();
            return null;
        }
        catch (Exception ex)
        {
            return ex;
        }
    }

    private static byte[] CreateWordDocument()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("Texto");
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static byte[] CreateExcelWorkbook()
    {
        using MemoryStream stream = new();
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["A1"].Value = "Texto";
            workbook.Save(stream);
        }

        return stream.ToArray();
    }
}
