using System;
using System.Linq;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.5a: the table model — creating tables, and reading and editing their rows, cells and
/// text. Formatting (widths, borders, shading, merge) is preserved but not modelled here.
/// </summary>
public sealed class WordTableTests
{
    [Fact]
    public void Tables_IsEmpty_WhenTheBodyHasNoTables()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("only paragraphs");

        Assert.Empty(document.Tables);
    }

    [Fact]
    public void AddTable_CreatesTheRequestedRowsAndColumns()
    {
        using WordDocument document = WordDocument.Create();

        Table table = document.AddTable(3, 4);

        Assert.Equal(3, table.RowCount);
        Assert.Equal(4, table.ColumnCount);
        Assert.Equal(3, table.Rows.Count);
        Assert.All(table.Rows, row => Assert.Equal(4, row.CellCount));
        Assert.Single(document.Tables);
    }

    [Fact]
    public void AddTable_GivesEveryCellOneEmptyParagraph()
    {
        using WordDocument document = WordDocument.Create();

        Table table = document.AddTable(2, 2);

        foreach (TableRow row in table.Rows)
        {
            foreach (TableCell cell in row.Cells)
            {
                Paragraph paragraph = Assert.Single(cell.Paragraphs);
                Assert.Equal(string.Empty, paragraph.Text);
                Assert.Equal(string.Empty, cell.Text);
            }
        }
    }

    [Fact]
    public void AddTable_AppendsTheTableAtTheEndOfTheBody()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("before");
        document.AddTable(1, 1);
        document.AddParagraph("after");

        byte[] produced = SaveBytes.ToBytes(document);

        // w:sectPr closes the body, as the schema requires; the table is the last content block.
        Assert.Equal(["p", "tbl", "p", "sectPr"], OoxmlValidation.ReadDocxBodyChildLocalNames(produced));
    }

    [Fact]
    public void AddTable_RejectsNonPositiveDimensions()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTable(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTable(-1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTable(1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddTable(1, -1));
    }

    [Fact]
    public void Tables_ThrowsWhenDisposed()
    {
        WordDocument document = WordDocument.Create();
        document.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = document.Tables);
        Assert.Throws<ObjectDisposedException>(() => document.AddTable(1, 1));
    }

    [Fact]
    public void Tables_ReturnsTheTablesInDocumentOrder()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            TableWithText("first"),
            new W.Paragraph(new W.Run(new W.Text("middle"))),
            TableWithText("second"));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(2, document.Tables.Count);
        Assert.Equal("first", document.Tables[0].Rows[0].Cells[0].Text);
        Assert.Equal("second", document.Tables[1].Rows[0].Cells[0].Text);
    }

    [Fact]
    public void Table_ColumnCount_IsZero_WhenTheGridIsMissing()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            new W.Table(new W.TableProperties(), new W.TableRow(new W.TableCell(new W.Paragraph()))));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(0, document.Tables[0].ColumnCount);
    }

    [Fact]
    public void Table_HasNoRows_WhenTheTableIsEmpty()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(
            new W.Table(new W.TableProperties(), new W.TableGrid(new W.GridColumn { Width = "2400" })));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(0, document.Tables[0].RowCount);
        Assert.Empty(document.Tables[0].Rows);
    }

    [Fact]
    public void TableRow_CellCount_ReportsTheStoredCells()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTable("before", "cell", "after");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        TableRow row = Assert.Single(document.Tables[0].Rows);

        Assert.Equal(1, row.CellCount);
        Assert.Single(row.Cells);
    }

    [Fact]
    public void Table_AddRow_AppendsARowWithOneCellPerGridColumn()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 3);

        TableRow row = table.AddRow();

        Assert.Equal(2, table.RowCount);
        Assert.Equal(3, row.CellCount);
        Assert.Equal(3, table.Rows[1].CellCount);
        Assert.Equal(2, document.Tables[0].RowCount);
    }

    [Fact]
    public void Table_AddRow_UsesTheStoredGridWhenTheTableComesFromAnotherTool()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTable("before", "cell", "after");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Table table = document.Tables[0];

        TableRow row = table.AddRow();

        Assert.Equal(1, row.CellCount);
        Assert.Equal(2, table.RowCount);
    }

    [Fact]
    public void TableCell_Text_ConcatenatesItsParagraphs()
    {
        W.TableCell cellElement = new(
            new W.Paragraph(new W.Run(new W.Text("first"))),
            new W.Paragraph());
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(TableAround(cellElement));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        TableCell cell = document.Tables[0].Rows[0].Cells[0];

        Assert.Equal(2, cell.Paragraphs.Count);
        Assert.Equal("first\n", cell.Text);
    }

    [Fact]
    public void TableCell_Text_Setter_ReplacesTheContentAndKeepsTheCellProperties()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithFormattedTable("original");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            TableCell cell = document.Tables[0].Rows[0].Cells[0];
            cell.Text = "replaced";

            Assert.Equal("replaced", cell.Text);
            Assert.Single(cell.Paragraphs);
            produced = SaveBytes.ToBytes(document);
        }

        string documentXml = TestPackage.ReadEntryText(produced, "word/document.xml");
        Assert.Contains("replaced", documentXml, StringComparison.Ordinal);
        Assert.DoesNotContain("original", documentXml, StringComparison.Ordinal);

        // The unmodelled cell properties must survive the edit.
        Assert.Contains("D9E2F3", documentXml, StringComparison.Ordinal);
        Assert.Contains("gridSpan", documentXml, StringComparison.Ordinal);
        Assert.Contains("w:tblW", documentXml, StringComparison.Ordinal);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void TableCell_Text_Setter_ToEmpty_LeavesOneEmptyParagraph()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTable("before", "cell", "after");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        TableCell cell = document.Tables[0].Rows[0].Cells[0];

        cell.Text = string.Empty;

        Assert.Equal(string.Empty, cell.Text);
        Paragraph paragraph = Assert.Single(cell.Paragraphs);
        Assert.Empty(paragraph.Runs);
    }

    [Fact]
    public void TableCell_Text_Setter_RejectsTabsAndLineBreaks()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        Assert.Throws<ArgumentException>(() => cell.Text = "a\tb");
        Assert.Throws<ArgumentException>(() => cell.Text = "a\nb");
    }

    [Fact]
    public void TableCell_Text_Setter_RejectsNull()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        Assert.Throws<ArgumentNullException>(() => cell.Text = null!);
    }

    [Fact]
    public void TableCell_AddParagraph_AppendsAParagraph()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        Paragraph added = cell.AddParagraph("second");

        Assert.Equal(2, cell.Paragraphs.Count);
        Assert.Equal("second", added.Text);
        Assert.Equal("\nsecond", cell.Text);
    }

    [Fact]
    public void TableCell_AddParagraph_RejectsTabsAndLineBreaks()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        Assert.Throws<ArgumentException>(() => cell.AddParagraph("a\tb"));
        Assert.Throws<ArgumentNullException>(() => cell.AddParagraph(null!));
    }

    [Fact]
    public void TableCell_Paragraphs_AreEditableThroughTheirOwnApi()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        cell.Paragraphs[0].Text = "edited";

        Assert.Equal("edited", cell.Text);
    }

    [Fact]
    public void Text_IncludesTableCellText()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTable("before", "TableCellText", "after");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal("before\nTableCellText\nafter", document.Text);
    }

    [Fact]
    public void Text_SeparatesCellsWithTabsAndRowsWithLineFeeds()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].Text = "a1";
        table.Rows[0].Cells[1].Text = "b1";
        table.Rows[1].Cells[0].Text = "a2";
        table.Rows[1].Cells[1].Text = "b2";

        Assert.Equal("a1\tb1\na2\tb2", document.Text);
    }

    [Fact]
    public void Text_PutsEveryParagraphOfACellOnItsOwnLine()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Text = "first";
        cell.AddParagraph("second");

        Assert.Equal("first\nsecond", document.Text);
    }

    [Fact]
    public void Text_DoesNotIncludeTextOfANestedTable()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithNestedTable("outer", "inner");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        // The nested table is preserved but not modelled, so its text is outside the scope of Text.
        Assert.Equal("outer", document.Text);
        Assert.Single(document.Tables);
        Assert.Single(document.Tables[0].Rows[0].Cells[0].Paragraphs);
    }

    [Fact]
    public void Text_DoesNotIncludeUnmodelledBodyContent()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithUnmodelledBodyContent("before");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal("before", document.Text);
    }

    [Fact]
    public void Types_ReportTheirSizeAsText()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 3);

        Assert.Equal("table (rows: 2, columns: 3)", table.ToString());
        Assert.Equal("table row (cells: 3)", table.Rows[0].ToString());
        Assert.Equal("table cell (paragraphs: 1)", table.Rows[0].Cells[0].ToString());
    }

    private static W.Table TableWithText(string text)
        => TableAround(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(text)))));

    private static W.Table TableAround(params W.TableCell[] cells)
        => new(
            new W.TableProperties(),
            new W.TableGrid(new W.GridColumn { Width = "2400" }),
            new W.TableRow(cells));
}
