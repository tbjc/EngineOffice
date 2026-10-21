using System;
using System.Linq;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.5c: interoperability of merged cells, verified with the Open XML SDK instead of the code
/// under test — produced packages are read back and validated independently.
/// </summary>
public sealed class WordTableMergeInteropTests
{
    [Fact]
    public void ProducedDocument_WithMergedCells_IsSchemaValid()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 3);
        table.Rows[0].Cells[0].Text = "one";
        table.Rows[0].Cells[1].Text = "two";

        // The vertical merge is applied first: the horizontal one shifts the cells beside it.
        table.MergeCellsVertically(0, 2, 2);
        table.Rows[0].MergeCells(0, 2);

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        W.Table read = Assert.Single(OoxmlValidation.ReadDocxTables(produced));
        W.TableRow firstRow = read.Elements<W.TableRow>().First();
        W.TableCell mergedCell = firstRow.Elements<W.TableCell>().First();
        Assert.Equal(2, mergedCell.TableCellProperties?.GridSpan?.Val?.Value);
        Assert.Equal(
            W.MergedCellValues.Restart,
            firstRow.Elements<W.TableCell>().Last().TableCellProperties?.VerticalMerge?.Val?.Value);
        Assert.Equal(
            W.MergedCellValues.Continue,
            read.Elements<W.TableRow>().Last().Elements<W.TableCell>().Last()
                .TableCellProperties?.VerticalMerge?.Val?.Value);

        // Every row still covers as many columns as the grid declares.
        int columns = read.GetFirstChild<W.TableGrid>()!.Elements<W.GridColumn>().Count();
        Assert.All(read.Elements<W.TableRow>(), row => Assert.Equal(columns, SpanSum(row)));
    }

    [Fact]
    public void ProducedDocument_EveryVerticalMergeState_IsSchemaValidAndReadable()
    {
        foreach (TableVerticalMerge state in Enum.GetValues<TableVerticalMerge>())
        {
            using WordDocument document = WordDocument.Create();
            document.AddTable(1, 1).Rows[0].Cells[0].VerticalMerge = state;

            byte[] produced = SaveBytes.ToBytes(document);
            W.MergedCellValues? stored = OoxmlValidation.ReadDocxFirstTableCellProperties(produced)
                ?.VerticalMerge?.Val?.Value;

            Assert.Empty(OoxmlValidation.ValidateDocx(produced));
            Assert.NotNull(stored);
        }
    }

    [Fact]
    public void ProducedDocument_WritesTheMergedRowWithTheFewerCellsOfAMergedRow()
    {
        using WordDocument document = WordDocument.Create();
        TableRow row = document.AddTable(1, 3).Rows[0];
        row.Cells[0].Text = "one";
        row.Cells[1].Text = "two";
        row.MergeCells(0, 2);

        W.TableRow read = OoxmlValidation.ReadDocxTables(SaveBytes.ToBytes(document))[0]
            .Elements<W.TableRow>().First();

        // A merged row holds one cell fewer per merged column, and the grid keeps its three.
        Assert.Equal(2, read.Elements<W.TableCell>().Count());
        Assert.Equal(3, SpanSum(read));
    }

    [Fact]
    public void OpenedDocument_WithMergeMarkupThisVersionDoesNotModel_KeepsItWhenMerging()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.Table(
            new W.TableProperties(
                new W.TableCellMarginDefault(new W.LeftMargin { Width = "108", Type = W.TableWidthUnitValues.Dxa }),
                new W.TableLook { Val = "04A0" }),
            new W.TableGrid(new W.GridColumn { Width = "2400" }, new W.GridColumn { Width = "2400" }),
            new W.TableRow(
                new W.TableCell(
                    new W.TableCellProperties(
                        new W.HorizontalMerge { Val = W.MergedCellValues.Restart },
                        new W.NoWrap()),
                    new W.Paragraph(new W.Run(new W.Text("one")))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text("two")))))));

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Tables[0].Rows[0].MergeCells(0, 2);
            produced = SaveBytes.ToBytes(document);
        }

        W.TableCellProperties properties = OoxmlValidation.ReadDocxFirstTableCellProperties(produced)!;
        Assert.NotNull(properties.HorizontalMerge);
        Assert.NotNull(properties.NoWrap);
        Assert.Equal("04A0", OoxmlValidation.ReadDocxFirstTableProperties(produced)?.TableLook?.Val);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void ProducedDocument_AfterSplittingAMergedCell_HasOneCellPerGridColumn()
    {
        byte[] produced;
        using (WordDocument document = WordDocument.Create())
        {
            TableRow row = document.AddTable(1, 2).Rows[0];
            row.Cells[0].Text = "one";
            row.Cells[1].Text = "two";
            row.MergeCells(0, 2);

            row.UnmergeCell(0);

            Assert.Equal(2, row.CellCount);
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        W.TableRow read = OoxmlValidation.ReadDocxTables(produced)[0].Elements<W.TableRow>().First();
        Assert.Equal(2, read.Elements<W.TableCell>().Count());
        Assert.All(read.Elements<W.TableCell>(), cell => Assert.Null(cell.TableCellProperties?.GridSpan));

        // The cell that was split keeps the paragraphs it had merged; the created one is empty.
        W.TableCell split = read.Elements<W.TableCell>().First();
        Assert.Equal(["one", "two"], split.Elements<W.Paragraph>().Select(paragraph => paragraph.InnerText));
        Assert.Empty(read.Elements<W.TableCell>().Last().Elements<W.Paragraph>().First().Elements<W.Run>());
    }

    private static int SpanSum(W.TableRow row)
        => row.Elements<W.TableCell>()
            .Sum(cell => cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
}
