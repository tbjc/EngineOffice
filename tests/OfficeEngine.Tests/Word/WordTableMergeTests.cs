using System;
using System.Linq;
using OfficeEngine;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.5c: merged cells. A cell states the columns it covers (<c>w:gridSpan</c>) and the part
/// it plays in a vertical merge (<c>w:vMerge</c>); merging and splitting set that markup and move or
/// create the cells it covers, never discarding content.
/// </summary>
public sealed class WordTableMergeTests
{
    [Fact]
    public void NewTable_CellsCoverOneColumnAndAreNotMerged()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(2, 2).Rows[0].Cells[0];

        Assert.Equal(1, cell.ColumnSpan);
        Assert.Null(cell.VerticalMerge);

        // A cell that is not merged states no merge markup at all.
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
    }

    [Theory]
    [InlineData(2, 2)]
    [InlineData(3, 3)]
    [InlineData(1, 1)]
    public void Cell_ColumnSpan_ReadsTheStoredSpan(int stored, int expected)
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithCellProperties(
            "text",
            new W.TableCellProperties(new W.GridSpan { Val = stored }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(expected, document.Tables[0].Rows[0].Cells[0].ColumnSpan);
    }

    [Fact]
    public void Cell_ColumnSpan_WithoutVal_IsRejected()
    {
        // ST_TcPrGridSpan requires w:val, so a stored span without it cannot be read.
        byte[] fixture = OoxmlValidation.CreateDocxWithCellProperties(
            "text",
            new W.TableCellProperties(new W.GridSpan()));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Tables[0].Rows[0].Cells[0].ColumnSpan);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-2)]
    [InlineData("many")]
    public void Cell_ColumnSpan_StoredValueThatIsNotAColumnCount_IsRejected(object stored)
    {
        W.GridSpan span = stored is int count
            ? new W.GridSpan { Val = count }
            : OoxmlValidation.WithAttribute(new W.GridSpan(), "val", (string)stored);
        byte[] fixture = OoxmlValidation.CreateDocxWithCellProperties("text", new W.TableCellProperties(span));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Tables[0].Rows[0].Cells[0].ColumnSpan);
    }

    [Fact]
    public void Cell_ColumnSpan_WritesTheStoredSpanInSchemaOrder()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(2, 2).Rows[0].Cells[0];
        cell.Width = TableWidth.Points(60.0);
        cell.VerticalAlignment = TableVerticalAlignment.Center;

        cell.ColumnSpan = 2;

        Assert.Equal(2, cell.ColumnSpan);

        // CT_TcPr order: tcW, gridSpan, vAlign.
        byte[] produced = SaveBytes.ToBytes(document);
        W.TableCellProperties properties = OoxmlValidation.ReadDocxFirstTableCellProperties(produced)!;
        Assert.Equal(
            ["tcW", "gridSpan", "vAlign"],
            properties.ChildElements.Select(child => child.LocalName));
        Assert.Equal(2, properties.GetFirstChild<W.GridSpan>()!.Val!.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Cell_ColumnSpan_SetToOne_RemovesTheStoredSpan()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.ColumnSpan = 2;

        cell.ColumnSpan = 1;

        // One column and no span are the same thing; w:gridSpan was the only property, so w:tcPr goes too.
        Assert.Equal(1, cell.ColumnSpan);
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Cell_ColumnSpan_OutOfRange_Throws(int span)
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => cell.ColumnSpan = span);

        // The rejected value left no markup behind.
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void Cell_VerticalMerge_ReadsTheStoredStates()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithVerticallyMergedTable("top", "middle", "bottom");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Table table = document.Tables[0];

        Assert.Equal(TableVerticalMerge.Restart, table.Rows[0].Cells[0].VerticalMerge);

        // ST_Merge: an omitted w:val means the cell continues the merge, and so does an explicit one.
        Assert.Equal(TableVerticalMerge.Continue, table.Rows[1].Cells[0].VerticalMerge);
        Assert.Equal(TableVerticalMerge.Continue, table.Rows[2].Cells[0].VerticalMerge);

        // The column beside the merge is not part of it.
        Assert.Null(table.Rows[0].Cells[1].VerticalMerge);
    }

    [Fact]
    public void Cell_VerticalMerge_UnknownValue_IsRejected()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithCellProperties(
            "text",
            new W.TableCellProperties(OoxmlValidation.WithAttribute(new W.VerticalMerge(), "val", "sideways")));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Tables[0].Rows[0].Cells[0].VerticalMerge);
    }

    [Fact]
    public void Cell_VerticalMerge_WritesTheStoredState()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        cell.VerticalMerge = TableVerticalMerge.Restart;
        byte[] restart = SaveBytes.ToBytes(document);
        Assert.Equal(
            "restart",
            OoxmlValidation.ReadDocxFirstTableCellProperties(restart)!
                .GetFirstChild<W.VerticalMerge>()!
                .Val!
                .InnerText);
        Assert.Empty(OoxmlValidation.ValidateDocx(restart));

        cell.VerticalMerge = TableVerticalMerge.Continue;
        byte[] continued = SaveBytes.ToBytes(document);

        // w:vMerge may omit w:val to mean "continue", but this version always states the value.
        Assert.Equal(
            "continue",
            OoxmlValidation.ReadDocxFirstTableCellProperties(continued)!
                .GetFirstChild<W.VerticalMerge>()!.Val!.InnerText);
        Assert.Empty(OoxmlValidation.ValidateDocx(continued));

        cell.VerticalMerge = null;
        Assert.Null(cell.VerticalMerge);
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void Cell_VerticalMerge_UndefinedValue_Throws()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => cell.VerticalMerge = (TableVerticalMerge)7);

        // The rejected value left no markup behind.
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void MergeCells_JoinsTheRangeIntoItsFirstCell()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 3).Rows[0];
        row.Cells[0].Text = "one";
        row.Cells[1].Text = "two";
        row.Cells[2].Text = "three";

        row.MergeCells(0, 2);

        Assert.Equal(2, row.CellCount);
        Assert.Equal(2, row.Cells[0].ColumnSpan);
        Assert.Equal(1, row.Cells[1].ColumnSpan);
        Assert.Equal("three", row.Cells[1].Text);
    }

    [Fact]
    public void MergeCells_MovesTheContentOfEveryCellInOrder()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 3).Rows[0];
        row.Cells[0].Text = "one";
        row.Cells[1].Text = "two";
        row.Cells[2].Text = "three";

        row.MergeCells(0, 3);

        // The paragraphs of the cells that are merged away arrive after the ones the cell already had.
        Assert.Equal(1, row.CellCount);
        Assert.Equal("one\ntwo\nthree", row.Cells[0].Text);
        Assert.Equal(3, row.Cells[0].ColumnSpan);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void MergeCells_AddsUpTheSpansOfAlreadyMergedCells()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 3).Rows[0];
        row.Cells[0].ColumnSpan = 2;

        row.MergeCells(0, 2);

        Assert.Equal(2, row.CellCount);
        Assert.Equal(3, row.Cells[0].ColumnSpan);
    }

    [Fact]
    public void MergeCells_AddsUpWidthsOfTheSameKind()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 3).Rows[0];
        row.Cells[0].Width = TableWidth.Points(120.0);
        row.Cells[1].Width = TableWidth.Points(60.0);
        row.Cells[2].Width = TableWidth.Percentage(10.0);

        row.MergeCells(0, 2);

        Assert.Equal(TableWidth.Points(180.0), row.Cells[0].Width);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void MergeCells_AddsUpPercentageWidths()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 2).Rows[0];
        row.Cells[0].Width = TableWidth.Percentage(25.0);
        row.Cells[1].Width = TableWidth.Percentage(25.0);

        row.MergeCells(0, 2);

        Assert.Equal(TableWidth.Percentage(50.0), row.Cells[0].Width);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MergeCells_WithWidthsThatCannotBeAddedUp_StatesNoWidth(bool mixedKinds)
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 2).Rows[0];
        row.Cells[0].Width = TableWidth.Points(120.0);
        row.Cells[1].Width = mixedKinds ? TableWidth.Percentage(25.0) : null;

        row.MergeCells(0, 2);

        // A width of each column of the grid is what the merged cell gets instead.
        Assert.Null(row.Cells[0].Width);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void MergeCells_OfACellOfAVerticalMerge_Throws()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 2).Rows[0];
        row.Cells[0].VerticalMerge = TableVerticalMerge.Restart;

        Assert.Throws<InvalidOperationException>(() => row.MergeCells(0, 2));

        // Nothing was merged.
        Assert.Equal(2, row.CellCount);
        Assert.Equal(1, row.Cells[0].ColumnSpan);
    }

    [Fact]
    public void MergeCells_OutOfRangeArguments_Throw()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 2).Rows[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => row.MergeCells(-1, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => row.MergeCells(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => row.MergeCells(1, 2));

        Assert.Equal(2, row.CellCount);
        Assert.Equal(1, row.Cells[0].ColumnSpan);
    }

    [Fact]
    public void MergeCells_KeepsTheNestedTableOfACellThatIsMergedAway()
    {
        W.TableCell withNestedTable = new(
            new W.Paragraph(new W.Run(new W.Text("outer"))),
            new W.Table(
                new W.TableProperties(),
                new W.TableGrid(new W.GridColumn { Width = "2400" }),
                new W.TableRow(new W.TableCell(new W.Paragraph(new W.Run(new W.Text("inner")))))));
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(TableWithCells(
            withNestedTable,
            new W.TableCell(new W.Paragraph(new W.Run(new W.Text("second"))))));

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Tables[0].Rows[0].MergeCells(0, 2);

            Assert.Equal("outer\nsecond", document.Tables[0].Rows[0].Cells[0].Text);
            Assert.Single(document.Tables[0].Rows[0].Cells);
            produced = SaveBytes.ToBytes(document);
        }

        // The nested table is content this version does not model, so it travels with the cell.
        Assert.Contains("inner", TestPackage.ReadEntryText(produced, "word/document.xml"), StringComparison.Ordinal);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void MergeCells_KeepsTheMarkupThisVersionDoesNotModel()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(TableWithCells(
            new W.TableCell(
                new W.TableCellProperties(new W.NoWrap()),
                new W.Paragraph(new W.Run(new W.Text("one")))),
            new W.TableCell(new W.Paragraph(new W.Run(new W.Text("two"))))));

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Tables[0].Rows[0].MergeCells(0, 2);
            produced = SaveBytes.ToBytes(document);
        }

        // The span of the kept cell is inserted before the markup this version does not model (CT_TcPr order).
        W.TableCellProperties properties = OoxmlValidation.ReadDocxFirstTableCellProperties(produced)!;
        Assert.Equal(["gridSpan", "noWrap"], properties.ChildElements.Select(child => child.LocalName));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void UnmergeCell_SplitsTheCellIntoOneCellPerColumn()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 3).Rows[0];
        row.Cells[0].Text = "one";
        row.Cells[1].Text = "two";
        row.Cells[2].Text = "three";
        row.MergeCells(0, 2);

        row.UnmergeCell(0);

        Assert.Equal(3, row.CellCount);
        Assert.Equal(1, row.Cells[0].ColumnSpan);
        Assert.Equal("one\ntwo", row.Cells[0].Text);
        Assert.Equal(string.Empty, row.Cells[1].Text);
        Assert.Null(row.Cells[1].Width);
        Assert.Equal("three", row.Cells[2].Text);
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void UnmergeCell_OfAStoredSpanOfOne_RemovesTheStoredSpan()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithCellProperties(
            "text",
            new W.TableCellProperties(new W.GridSpan { Val = 1 }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        TableRow row = document.Tables[0].Rows[0];

        row.UnmergeCell(0);

        Assert.Equal(1, row.CellCount);
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void UnmergeCell_OfACellCoveringOneColumn_ChangesNothing()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 3).Rows[0];

        row.UnmergeCell(1);

        Assert.Equal(3, row.CellCount);
        Assert.Equal(1, row.Cells[1].ColumnSpan);
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void UnmergeCell_OutOfRangeArguments_Throw()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 2).Rows[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => row.UnmergeCell(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => row.UnmergeCell(2));

        Assert.Equal(2, row.CellCount);
    }

    [Fact]
    public void UnmergeCell_OfACellOfAVerticalMerge_Throws()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 2).Rows[0];
        row.Cells[0].ColumnSpan = 2;
        row.Cells[0].VerticalMerge = TableVerticalMerge.Restart;

        Assert.Throws<InvalidOperationException>(() => row.UnmergeCell(0));

        Assert.Equal(2, row.CellCount);
        Assert.Equal(2, row.Cells[0].ColumnSpan);
    }

    [Fact]
    public void MergeCellsVertically_MarksTheTopCellAndTheRest()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(3, 2);

        table.MergeCellsVertically(0, 1, 3);

        Assert.Equal(TableVerticalMerge.Restart, table.Rows[0].Cells[1].VerticalMerge);
        Assert.Equal(TableVerticalMerge.Continue, table.Rows[1].Cells[1].VerticalMerge);
        Assert.Equal(TableVerticalMerge.Continue, table.Rows[2].Cells[1].VerticalMerge);

        // The other column is not part of the merge, and no cell was added or removed.
        Assert.Null(table.Rows[0].Cells[0].VerticalMerge);
        Assert.Equal(3, table.RowCount);
        Assert.Equal(2, table.Rows[0].CellCount);
    }

    [Fact]
    public void MergeCellsVertically_LeavesTheContentWhereItIs()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 1);
        table.Rows[0].Cells[0].Text = "top";
        table.Rows[1].Cells[0].Text = "bottom";

        table.MergeCellsVertically(0, 0, 2);

        // A vertical merge draws one cell, but the content of the cells below it is not moved.
        Assert.Equal("top", table.Rows[0].Cells[0].Text);
        Assert.Equal("bottom", table.Rows[1].Cells[0].Text);
    }

    [Fact]
    public void MergeCellsVertically_Twice_KeepsTheSameMarks()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 1);
        table.MergeCellsVertically(0, 0, 2);

        table.MergeCellsVertically(0, 0, 2);

        Assert.Equal(TableVerticalMerge.Restart, table.Rows[0].Cells[0].VerticalMerge);
        Assert.Equal(TableVerticalMerge.Continue, table.Rows[1].Cells[0].VerticalMerge);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.Single(OoxmlValidation.ReadDocxTables(produced));
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void MergeCellsVertically_WhenTheTopCellContinuesAMergeAbove_Throws()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(3, 1);
        table.MergeCellsVertically(0, 0, 3);

        Assert.Throws<InvalidOperationException>(() => table.MergeCellsVertically(1, 0, 2));

        // The existing merge is left as it was.
        Assert.Equal(TableVerticalMerge.Restart, table.Rows[0].Cells[0].VerticalMerge);
        Assert.Equal(TableVerticalMerge.Continue, table.Rows[1].Cells[0].VerticalMerge);
    }

    [Fact]
    public void MergeCellsVertically_WhenTheCellsCoverDifferentSpans_Throws()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].ColumnSpan = 2;

        Assert.Throws<InvalidOperationException>(() => table.MergeCellsVertically(0, 0, 2));

        Assert.Null(table.Rows[0].Cells[0].VerticalMerge);
        Assert.Null(table.Rows[1].Cells[0].VerticalMerge);
    }

    [Fact]
    public void MergeCellsVertically_OutOfRangeArguments_Throw()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 2);

        Assert.Throws<ArgumentOutOfRangeException>(() => table.MergeCellsVertically(-1, 0, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.MergeCellsVertically(0, -1, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.MergeCellsVertically(0, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.MergeCellsVertically(1, 0, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.MergeCellsVertically(0, 2, 2));

        Assert.Null(table.Rows[0].Cells[0].VerticalMerge);
    }

    [Fact]
    public void UnmergeCellsVertically_ClearsTheMarksOfTheRange()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithVerticallyMergedTable("top", "middle", "bottom");

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Table table = document.Tables[0];
            table.UnmergeCellsVertically(0, 0, 3);

            Assert.Null(table.Rows[0].Cells[0].VerticalMerge);
            Assert.Null(table.Rows[1].Cells[0].VerticalMerge);
            Assert.Null(table.Rows[2].Cells[0].VerticalMerge);
            produced = SaveBytes.ToBytes(document);
        }

        // The cells keep their content, and their properties carried no other markup to keep.
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(produced));
        Assert.Contains("top", TestPackage.ReadEntryText(produced, "word/document.xml"), StringComparison.Ordinal);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void UnmergeCellsVertically_OutOfRangeArguments_Throw()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 2);
        table.MergeCellsVertically(0, 0, 2);

        Assert.Throws<ArgumentOutOfRangeException>(() => table.UnmergeCellsVertically(-1, 0, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.UnmergeCellsVertically(0, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.UnmergeCellsVertically(1, 0, 2));

        // The merge is left as it was.
        Assert.Equal(TableVerticalMerge.Restart, table.Rows[0].Cells[0].VerticalMerge);
        Assert.Equal(TableVerticalMerge.Continue, table.Rows[1].Cells[0].VerticalMerge);
    }

    [Fact]
    public void MergedTable_SurvivesASaveAndReopen()
    {
        byte[] produced;
        using (WordDocument document = WordDocument.Create())
        {
            Table table = document.AddTable(2, 3);
            table.Rows[0].Cells[0].Text = "one";
            table.Rows[0].Cells[1].Text = "two";
            table.MergeCellsVertically(0, 2, 2);
            table.Rows[0].MergeCells(0, 2);
            produced = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Table read = reopened.Tables[0];

        Assert.Equal(2, read.Rows[0].Cells[0].ColumnSpan);
        Assert.Equal("one\ntwo", read.Rows[0].Cells[0].Text);
        Assert.Equal(TableVerticalMerge.Restart, read.Rows[0].Cells[1].VerticalMerge);
        Assert.Equal(1, read.Rows[1].Cells[0].ColumnSpan);
        Assert.Equal(TableVerticalMerge.Continue, read.Rows[1].Cells[2].VerticalMerge);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    private static W.Table TableWithCells(params W.TableCell[] cells)
        => new(
            new W.TableProperties(),
            new W.TableGrid(cells.Select(_ => new W.GridColumn { Width = "2400" })),
            new W.TableRow(cells));
}
