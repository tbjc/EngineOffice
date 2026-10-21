using DocumentFormat.OpenXml.Validation;
using OfficeEngine.Diagnostics;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Tests that verify produced and consumed packages with the Open XML SDK instead of the
/// code under test. These are the Phase 1 interoperability checks.
/// </summary>
public sealed class WordInteropTests
{
    [Fact]
    public void ProducedPackage_IsAValidWordprocessingDocument()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Validación con el SDK");
        byte[] bytes = SaveBytes.ToBytes(document);

        IReadOnlyList<ValidationErrorInfo> errors = OoxmlValidation.ValidateDocx(bytes);

        Assert.Empty(errors);
    }

    [Fact]
    public void ProducedPackage_ContainsExpectedContentTypeAndDocumentXml()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Contenido verificable");
        byte[] bytes = SaveBytes.ToBytes(document);

        Assert.True(TestPackage.HasEntry(bytes, "[Content_Types].xml"));
        Assert.True(TestPackage.HasEntry(bytes, "word/document.xml"));

        string documentXml = TestPackage.ReadEntryText(bytes, "word/document.xml");
        Assert.Contains("Contenido verificable", documentXml, StringComparison.Ordinal);
        Assert.Contains("w:document", documentXml, StringComparison.Ordinal);
        Assert.Contains("w:body", documentXml, StringComparison.Ordinal);
    }

    [Fact]
    public void ProducedPackage_IsReadableByTheOpenXmlSdk()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Hola SDK");
        byte[] bytes = SaveBytes.ToBytes(document);

        Assert.Equal("Hola SDK", OoxmlValidation.ReadDocxText(bytes));
    }

    [Fact]
    public void OpenedDocument_WithTable_ExposesItAndKeepsItWhenTheBodyIsEdited()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTable("before", "TableCellText", "after");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            // A modelled table is no longer reported as unmodelled content.
            Assert.Empty(document.Diagnostics);

            // A table sits at the same level as a paragraph, so the two paragraphs stay visible.
            Assert.Equal(2, document.Paragraphs.Count);
            Assert.Equal("before\nTableCellText\nafter", document.Text);

            Table table = Assert.Single(document.Tables);
            Assert.Equal("TableCellText", table.Rows[0].Cells[0].Text);

            document.AddParagraph("appended");
            produced = SaveBytes.ToBytes(document);
        }

        // The table must still be present after a round-trip.
        string documentXml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("w:tbl", documentXml, StringComparison.Ordinal);
        Assert.Contains("TableCellText", documentXml, StringComparison.Ordinal);
        Assert.Contains("appended", documentXml, StringComparison.Ordinal);

        // The result must remain a valid document that the SDK can read.
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal("beforeTableCellTextafterappended", OoxmlValidation.ReadDocxText(produced));
    }

    [Fact]
    public void OpenedDocument_WithUnmodelledBodyContent_ReportsDiagnosticAndPreservesIt()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithUnmodelledBodyContent("before");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            // The content control is not modelled, so it must be reported, not dropped.
            OfficeEngineDiagnostic diagnostic = Assert.Single(document.Diagnostics);
            Assert.Equal(DiagnosticSeverity.Info, diagnostic.Severity);
            Assert.Equal(WordDiagnosticCodesForTest.Unmodelled, diagnostic.Code);
            Assert.Contains("sdt", diagnostic.Message, StringComparison.OrdinalIgnoreCase);

            // Only the top-level paragraph is exposed.
            Assert.Single(document.Paragraphs);
            Assert.Equal("before", document.Text);

            document.AddParagraph("appended");
            produced = SaveBytes.ToBytes(document);
        }

        // The unknown content must still be present after a round-trip.
        string documentXml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("w:sdt", documentXml, StringComparison.Ordinal);
        Assert.Contains("appended", documentXml, StringComparison.Ordinal);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void OpenedDocument_WithModelledContent_HasNoDiagnostics()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTable("only", "table", "content");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(document.Diagnostics);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }
}

/// <summary>
/// Mirrors the internal diagnostic code so the test can assert on it without widening the public API.
/// </summary>
internal static class WordDiagnosticCodesForTest
{
    internal const string Unmodelled = "WORD-0001-UNMODELLED-BODY-CONTENT";
}
