using System;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.7: headers and footers per section. A section references header and footer parts by type,
/// several sections can share a part, and a value that cannot be read fails with a typed error instead of
/// being guessed.
/// </summary>
public sealed class WordHeaderFooterTests
{
    [Fact]
    public void Create_HasNoHeadersOrFooters()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];

        foreach (HeaderFooterType type in Enum.GetValues<HeaderFooterType>())
        {
            Assert.Null(section.Header(type));
            Assert.Null(section.Footer(type));
        }

        Assert.False(section.DifferentFirstPage);
    }

    [Fact]
    public void AddHeader_CreatesAnEmptyHeaderAndWritesAReference()
    {
        using WordDocument document = WordDocument.Create();

        HeaderFooter header = document.Sections[0].AddHeader(HeaderFooterType.Primary);

        Assert.True(header.IsHeader);
        Assert.Single(header.Paragraphs);
        Assert.Equal(string.Empty, header.Text);
        Assert.NotNull(document.Sections[0].Header(HeaderFooterType.Primary));

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        W.SectionProperties stored = OoxmlValidation.ReadDocxFinalSectionProperties(produced)!;
        W.HeaderReference reference = Assert.Single(stored.Elements<W.HeaderReference>());
        Assert.Equal(W.HeaderFooterValues.Default, reference.Type!.Value);
        Assert.Equal(1, OoxmlValidation.CountDocxHeaderParts(produced));
    }

    [Fact]
    public void Header_CanBeEditedThroughItsParagraphs()
    {
        using WordDocument document = WordDocument.Create();
        HeaderFooter header = document.Sections[0].AddHeader(HeaderFooterType.Primary);

        header.Paragraphs[0].Text = "Título";
        header.AddParagraph("Subtítulo");

        Assert.Equal("Título\nSubtítulo", header.Text);

        byte[] produced = SaveBytes.ToBytes(document);
        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        HeaderFooter? reopenedHeader = reopened.Sections[0].Header(HeaderFooterType.Primary);

        Assert.NotNull(reopenedHeader);
        Assert.Equal("Título\nSubtítulo", reopenedHeader!.Text);
        Assert.Equal(2, reopenedHeader.Paragraphs.Count);
    }

    [Fact]
    public void Footer_CanHoldATable()
    {
        using WordDocument document = WordDocument.Create();
        HeaderFooter footer = document.Sections[0].AddFooter(HeaderFooterType.Primary);

        Table table = footer.AddTable(1, 2);
        table.Rows[0].Cells[0].Text = "Página";
        table.Rows[0].Cells[1].Text = "1";

        Assert.False(footer.IsHeader);
        Assert.Single(footer.Tables);

        // A new footer starts with an empty paragraph, so the table is the second block.
        Assert.Equal("\nPágina\t1", footer.Text);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void AddHeader_IsIdempotent()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];

        HeaderFooter first = section.AddHeader(HeaderFooterType.Primary);
        first.Paragraphs[0].Text = "Texto";
        HeaderFooter second = section.AddHeader(HeaderFooterType.Primary);

        Assert.Equal("Texto", second.Text);
        W.SectionProperties stored = OoxmlValidation.ReadDocxFinalSectionProperties(SaveBytes.ToBytes(document))!;
        Assert.Single(stored.Elements<W.HeaderReference>());
    }

    [Fact]
    public void EveryType_CanBeAddedAndReadBack()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];

        foreach (HeaderFooterType type in Enum.GetValues<HeaderFooterType>())
        {
            section.AddHeader(type).Paragraphs[0].Text = $"cabecera {type}";
            section.AddFooter(type).Paragraphs[0].Text = $"pie {type}";
        }

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(3, OoxmlValidation.CountDocxHeaderParts(produced));
        Assert.Equal(3, OoxmlValidation.CountDocxFooterParts(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        foreach (HeaderFooterType type in Enum.GetValues<HeaderFooterType>())
        {
            Assert.Equal($"cabecera {type}", reopened.Sections[0].Header(type)!.Text);
            Assert.Equal($"pie {type}", reopened.Sections[0].Footer(type)!.Text);
        }
    }

    [Fact]
    public void HeaderAndFooter_RoundTripThroughAFile()
    {
        string directory = TestData.NewTemporaryDirectory();
        string path = Path.Combine(directory, "headers.docx");
        try
        {
            using (WordDocument document = WordDocument.Create())
            {
                document.Sections[0].AddHeader(HeaderFooterType.Primary).Paragraphs[0].Text = "Cabecera";
                document.Sections[0].AddFooter(HeaderFooterType.Primary).Paragraphs[0].Text = "Pie";
                document.Save(path);
            }

            using WordDocument reopened = WordDocument.Open(path);

            Assert.Equal("Cabecera", reopened.Sections[0].Header(HeaderFooterType.Primary)!.Text);
            Assert.Equal("Pie", reopened.Sections[0].Footer(HeaderFooterType.Primary)!.Text);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void DifferentFirstPage_RoundTripsAndClears()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];

        section.DifferentFirstPage = true;
        Assert.True(section.DifferentFirstPage);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.NotNull(OoxmlValidation.ReadDocxFinalSectionProperties(produced)?.GetFirstChild<W.TitlePage>());
        using (WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced)))
        {
            Assert.True(reopened.Sections[0].DifferentFirstPage);
        }

        section.DifferentFirstPage = false;
        Assert.False(section.DifferentFirstPage);
        Assert.Null(OoxmlValidation.ReadDocxFinalSectionProperties(SaveBytes.ToBytes(document))?.GetFirstChild<W.TitlePage>());
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void RemoveHeader_RemovesTheReferenceAndKeepsThePart()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];
        section.AddHeader(HeaderFooterType.Primary).Paragraphs[0].Text = "Texto";

        Assert.True(section.RemoveHeader(HeaderFooterType.Primary));
        Assert.Null(section.Header(HeaderFooterType.Primary));
        Assert.False(section.RemoveHeader(HeaderFooterType.Primary));

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxHeaderParts(produced));
        Assert.Empty(OoxmlValidation.ReadDocxFinalSectionProperties(produced)!.Elements<W.HeaderReference>());
    }

    [Fact]
    public void AddSection_NewSectionSharesTheHeaderOfThePreviousOne()
    {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].AddHeader(HeaderFooterType.Primary).Paragraphs[0].Text = "Compartido";

        Section second = document.AddSection();

        Assert.Equal("Compartido", second.Header(HeaderFooterType.Primary)!.Text);

        // The part is shared: editing it through the new section changes what the previous one shows.
        second.Header(HeaderFooterType.Primary)!.Paragraphs[0].Text = "Cambiado";
        Assert.Equal("Cambiado", document.Sections[0].Header(HeaderFooterType.Primary)!.Text);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(1, OoxmlValidation.CountDocxHeaderParts(produced));
    }

    [Fact]
    public void AddHeader_OnDifferentSections_KeepsThemIndependent()
    {
        using WordDocument document = WordDocument.Create();
        Section first = document.Sections[0];
        Section second = document.AddSection();

        first.AddHeader(HeaderFooterType.Primary).Paragraphs[0].Text = "Primera";
        second.AddHeader(HeaderFooterType.Primary).Paragraphs[0].Text = "Segunda";

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal(2, OoxmlValidation.CountDocxHeaderParts(produced));

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal("Primera", reopened.Sections[0].Header(HeaderFooterType.Primary)!.Text);
        Assert.Equal("Segunda", reopened.Sections[1].Header(HeaderFooterType.Primary)!.Text);
    }

    [Fact]
    public void Header_RejectsAnUndefinedType()
    {
        using WordDocument document = WordDocument.Create();
        Section section = document.Sections[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => section.Header((HeaderFooterType)99));
        Assert.Throws<ArgumentOutOfRangeException>(() => section.AddFooter((HeaderFooterType)99));
        Assert.Throws<ArgumentOutOfRangeException>(() => section.RemoveHeader((HeaderFooterType)99));
    }

    [Fact]
    public void Header_ReportsAReferenceWithoutType()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            new W.Paragraph(new W.Run(new W.Text("texto"))),
            new W.SectionProperties(new W.HeaderReference { Id = "rId1" }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].Header(HeaderFooterType.Primary));
    }

    [Fact]
    public void Header_ReportsAReferenceToAMissingPart()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            new W.Paragraph(new W.Run(new W.Text("texto"))),
            new W.SectionProperties(new W.HeaderReference
            {
                Type = W.HeaderFooterValues.Default,
                Id = "rId99",
            }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].Header(HeaderFooterType.Primary));
    }

    [Fact]
    public void Header_ReportsAMalformedType()
    {
        W.HeaderReference reference = new() { Id = "rId1" };
        OoxmlValidation.WithAttribute(reference, "type", "diagonal");

        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            new W.Paragraph(new W.Run(new W.Text("texto"))),
            new W.SectionProperties(reference));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].Header(HeaderFooterType.Primary));
    }

    [Fact]
    public void DifferentFirstPage_ReportsAMalformedValue()
    {
        W.TitlePage titlePage = new();
        OoxmlValidation.WithAttribute(titlePage, "val", "maybe");

        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            new W.Paragraph(new W.Run(new W.Text("texto"))),
            new W.SectionProperties(titlePage));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Sections[0].DifferentFirstPage);
    }

    [Fact]
    public void HeaderFooter_ValidatesTheContentItAdds()
    {
        using WordDocument document = WordDocument.Create();
        HeaderFooter header = document.Sections[0].AddHeader(HeaderFooterType.Primary);

        Assert.Throws<ArgumentNullException>(() => header.AddParagraph(null!));
        Assert.Throws<ArgumentException>(() => header.AddParagraph("con\ttabulador"));
        Assert.Throws<ArgumentOutOfRangeException>(() => header.AddTable(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => header.AddTable(1, 0));
    }
}
