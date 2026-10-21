using System;
using System.Linq;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.5a: interoperability of tables, verified with the Open XML SDK instead of the code
/// under test — produced packages are read back and validated independently.
/// </summary>
public sealed class WordTableInteropTests
{
    [Fact]
    public void ProducedDocument_WithTable_IsSchemaValid()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 3);
        table.Rows[0].Cells[0].Text = "a1";
        table.AddRow();

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void ProducedDocument_WritesTheTableChildrenInSchemaOrder()
    {
        using WordDocument document = WordDocument.Create();
        document.AddTable(2, 2);

        byte[] produced = SaveBytes.ToBytes(document);
        W.Table table = Assert.Single(OoxmlValidation.ReadDocxTables(produced));

        Assert.Equal(
            ["tblPr", "tblGrid", "tr", "tr"],
            table.ChildElements.Select(child => child.LocalName));
    }

    [Fact]
    public void ProducedDocument_DeclaresOneGridColumnPerColumn()
    {
        using WordDocument document = WordDocument.Create();
        document.AddTable(1, 3);

        byte[] produced = SaveBytes.ToBytes(document);
        W.TableGrid grid = Assert.IsType<W.TableGrid>(OoxmlValidation.ReadDocxTables(produced)[0].GetFirstChild<W.TableGrid>());
        IReadOnlyList<W.GridColumn> columns = grid.Elements<W.GridColumn>().ToList();

        Assert.Equal(3, columns.Count);
        Assert.All(columns, column => Assert.Equal("2400", column.Width?.Value));
    }

    [Fact]
    public void ProducedDocument_TableText_IsReadableByTheOpenXmlSdk()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].Text = "a1";
        table.Rows[0].Cells[1].Text = "b1";
        table.Rows[1].Cells[0].Text = "a2";
        table.Rows[1].Cells[1].Text = "b2";

        byte[] produced = SaveBytes.ToBytes(document);
        W.Table read = Assert.Single(OoxmlValidation.ReadDocxTables(produced));

        Assert.Equal(2, read.Elements<W.TableRow>().Count());
        Assert.Equal("a1b1a2b2", OoxmlValidation.ReadDocxText(produced));
        Assert.Equal("a1", read.Elements<W.TableRow>().First().Elements<W.TableCell>().First().InnerText);
    }

    [Fact]
    public void ProducedDocument_InterleavesParagraphsAndTables()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("before");
        document.AddTable(1, 1);
        document.AddParagraph("after");

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal(["p", "tbl", "p", "sectPr"], OoxmlValidation.ReadDocxBodyChildLocalNames(produced));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void ProducedDocument_Table_KeepsItsCellTextThroughRepeatedSaves()
    {
        using WordDocument document = WordDocument.Create();
        document.AddTable(1, 2).Rows[0].Cells[0].Text = "persisted";

        byte[] first = SaveBytes.ToBytes(document);
        byte[] second = SaveBytes.ToBytes(document);

        Assert.Equal("persisted", OoxmlValidation.ReadDocxText(first));
        Assert.Equal("persisted", OoxmlValidation.ReadDocxText(second));
        Assert.Empty(OoxmlValidation.ValidateDocx(second));
    }

    [Fact]
    public void OpenedDocument_WithFormattedTable_KeepsUnmodelledMarkupWhenTheContentIsEdited()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithFormattedTable("original");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Tables[0].Rows[0].Cells[0].Text = "edited";
            produced = SaveBytes.ToBytes(document);
        }

        W.Table table = Assert.Single(OoxmlValidation.ReadDocxTables(produced));
        W.TableCell cell = table.Elements<W.TableRow>().First().Elements<W.TableCell>().First();
        W.TableCellProperties properties = Assert.IsType<W.TableCellProperties>(cell.GetFirstChild<W.TableCellProperties>());

        Assert.Equal("edited", cell.InnerText);
        Assert.Equal(2, properties.GetFirstChild<W.GridSpan>()?.Val?.Value);
        Assert.Equal("D9E2F3", properties.GetFirstChild<W.Shading>()?.Fill?.Value);
        Assert.NotNull(table.GetFirstChild<W.TableProperties>()?.GetFirstChild<W.TableWidth>());
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void OpenedDocument_WithNestedTable_PreservesTheNestedTableWhenTheOuterContentIsEdited()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithNestedTable("outer", "inner");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Tables[0].Rows[0].Cells[0].Text = "replaced";
            produced = SaveBytes.ToBytes(document);
        }

        string documentXml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("replaced", documentXml, StringComparison.Ordinal);
        Assert.DoesNotContain("outer", documentXml, StringComparison.Ordinal);
        Assert.Contains("inner", documentXml, StringComparison.Ordinal);

        // The nested table is preserved, so the produced document still holds two tables in total.
        Assert.Equal(2, CountTables(documentXml));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    private static int CountTables(string documentXml)
        => documentXml.Split("<w:tbl>", StringSplitOptions.None).Length - 1;
}
