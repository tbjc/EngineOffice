using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// A round-trip over generated documents: many different shapes, always the same result. It is the
/// regression net of the model: a change that loses or reorders content shows up here.
/// </summary>
public sealed class WordRoundTripRegressionTests
{
    [Fact]
    public void GeneratedDocuments_RoundTripWithoutLosingText()
    {
        for (uint seed = 1; seed <= 12; seed++)
        {
            DocumentGenerator generator = new(seed);
            string expected;
            byte[] package;

            using (WordDocument document = generator.CreateDocument())
            {
                expected = document.Text;
                using MemoryStream stream = new();
                document.Save(stream);
                package = stream.ToArray();
            }

            using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(package));
            Assert.Equal(expected, reopened.Text);
        }
    }

    [Fact]
    public void GeneratedDocuments_KeepTheirStructure()
    {
        DocumentGenerator generator = new(99);
        int paragraphs;
        int tables;
        byte[] package;

        using (WordDocument document = generator.CreateDocument())
        {
            paragraphs = document.Paragraphs.Count;
            tables = document.Tables.Count;
            using MemoryStream stream = new();
            document.Save(stream);
            package = stream.ToArray();
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(package));
        Assert.Equal(paragraphs, reopened.Paragraphs.Count);
        Assert.Equal(tables, reopened.Tables.Count);
    }
}
