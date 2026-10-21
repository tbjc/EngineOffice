using System;
using System.Linq;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.5b: interoperability of table and cell formatting, verified with the Open XML SDK
/// instead of the code under test — produced packages are read back and validated independently.
/// </summary>
public sealed class WordTableFormattingInteropTests
{
    [Fact]
    public void ProducedDocument_WithTableAndCellFormatting_IsSchemaValid()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(2, 2);
        TableCell cell = table.Rows[0].Cells[0];

        table.Width = TableWidth.Percentage(80.0);
        table.Alignment = TableAlignment.Center;
        table.Layout = TableLayout.Fixed;
        table.ShadingColor = new RgbColor(0xEE, 0xEE, 0xEE);
        table.BorderTop = new TableBorder(BorderStyle.SingleLine, 1.5);
        table.BorderInsideVertical = new TableBorder(BorderStyle.Dotted, 0.5);

        cell.Width = TableWidth.Points(120.0);
        cell.VerticalAlignment = TableVerticalAlignment.Center;
        cell.ShadingColor = new RgbColor(0xD9, 0xE2, 0xF3);
        cell.BorderBottom = new TableBorder(BorderStyle.DoubleLine, 0.75, new RgbColor(0x00, 0x00, 0xFF));

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        W.Table read = Assert.Single(OoxmlValidation.ReadDocxTables(produced));
        W.TableProperties properties = read.TableProperties!;
        Assert.Equal("4000", properties.TableWidth?.Width?.Value);
        Assert.Equal("pct", properties.TableWidth?.Type?.InnerText);
        Assert.Equal(W.TableRowAlignmentValues.Center, properties.TableJustification?.Val?.Value);
        Assert.Equal(W.TableLayoutValues.Fixed, properties.TableLayout?.Type?.Value);
        Assert.Equal(W.BorderValues.Single, properties.TableBorders?.TopBorder?.Val?.Value);
        Assert.Equal(W.BorderValues.Dotted, properties.TableBorders?.InsideVerticalBorder?.Val?.Value);

        W.TableCell cellElement = read.Elements<W.TableRow>().First().Elements<W.TableCell>().First();
        W.TableCellProperties cellProperties = cellElement.TableCellProperties!;
        Assert.Equal("2400", cellProperties.TableCellWidth?.Width?.Value);
        Assert.Equal("dxa", cellProperties.TableCellWidth?.Type?.InnerText);
        Assert.Equal(W.TableVerticalAlignmentValues.Center, cellProperties.TableCellVerticalAlignment?.Val?.Value);
        Assert.Equal("D9E2F3", cellProperties.Shading?.Fill?.Value);
        Assert.Equal(W.BorderValues.Double, cellProperties.TableCellBorders?.BottomBorder?.Val?.Value);
        Assert.Equal(6U, cellProperties.TableCellBorders?.BottomBorder?.Size?.Value);
        Assert.Equal("0000FF", cellProperties.TableCellBorders?.BottomBorder?.Color?.Value);
    }

    [Fact]
    public void ProducedDocument_WritesTablePropertiesInSchemaOrder()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);
        table.Width = TableWidth.Points(200.0);
        table.Alignment = TableAlignment.Left;
        table.BorderTop = new TableBorder(BorderStyle.SingleLine, 0.5);
        table.ShadingColor = new RgbColor(0xFF, 0x00, 0x00);
        table.Layout = TableLayout.AutoFit;

        W.TableProperties properties = OoxmlValidation.ReadDocxFirstTableProperties(SaveBytes.ToBytes(document))!;

        Assert.Equal(
            ["tblW", "jc", "tblBorders", "shd", "tblLayout"],
            properties.ChildElements.Select(child => child.LocalName));
    }

    [Fact]
    public void ProducedDocument_WritesCellPropertiesInSchemaOrder()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Width = TableWidth.Points(60.0);
        cell.BorderLeft = new TableBorder(BorderStyle.SingleLine, 0.5);
        cell.ShadingColor = new RgbColor(0x00, 0xFF, 0x00);
        cell.VerticalAlignment = TableVerticalAlignment.Bottom;

        W.TableCellProperties properties =
            OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document))!;

        Assert.Equal(
            ["tcW", "tcBorders", "shd", "vAlign"],
            properties.ChildElements.Select(child => child.LocalName));
    }

    [Fact]
    public void ProducedDocument_EveryBorderStyle_IsSchemaValidAndReadable()
    {
        foreach (BorderStyle style in Enum.GetValues<BorderStyle>())
        {
            using WordDocument document = WordDocument.Create();
            document.AddTable(1, 1).BorderTop = new TableBorder(style, 0.5);

            byte[] produced = SaveBytes.ToBytes(document);
            W.BorderValues? stored = OoxmlValidation.ReadDocxFirstTableProperties(produced)
                ?.TableBorders?.TopBorder?.Val?.Value;

            Assert.Empty(OoxmlValidation.ValidateDocx(produced));
            Assert.NotNull(stored);
        }
    }

    [Fact]
    public void OpenedDocument_WithFormattedTable_KeepsUnmodelledMarkupWhenFormattingIsEdited()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.Table(
            new W.TableProperties(
                new W.TableStyle { Val = "GridTable4Accent1" },
                new W.TableWidth { Width = "0", Type = W.TableWidthUnitValues.Auto },
                new W.TableCellMarginDefault(new W.TopMargin { Width = "72", Type = W.TableWidthUnitValues.Dxa }),
                new W.TableLook { Val = "04A0" }),
            new W.TableGrid(new W.GridColumn { Width = "2400" }),
            new W.TableRow(new W.TableCell(
                new W.TableCellProperties(
                    new W.TableCellBorders(new W.InsideHorizontalBorder { Val = W.BorderValues.Dashed }),
                    new W.TableCellMargin(new W.LeftMargin { Width = "108", Type = W.TableWidthUnitValues.Dxa })),
                new W.Paragraph(new W.Run(new W.Text("cell")))))));

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Table table = document.Tables[0];
            table.Width = TableWidth.Points(250.0);
            table.BorderBottom = new TableBorder(BorderStyle.SingleLine, 0.5);
            table.Rows[0].Cells[0].ShadingColor = new RgbColor(0xAB, 0xCD, 0xEF);
            produced = SaveBytes.ToBytes(document);
        }

        W.TableProperties properties = OoxmlValidation.ReadDocxFirstTableProperties(produced)!;
        Assert.Equal("GridTable4Accent1", properties.TableStyle?.Val);
        Assert.NotNull(properties.TableCellMarginDefault);
        Assert.Equal("04A0", properties.TableLook?.Val);

        W.TableCellProperties cellProperties = OoxmlValidation.ReadDocxFirstTableCellProperties(produced)!;
        Assert.Equal(W.BorderValues.Dashed, cellProperties.TableCellBorders?.InsideHorizontalBorder?.Val?.Value);
        Assert.NotNull(cellProperties.TableCellMargin);

        W.TableWidth? width = OoxmlValidation.ReadDocxFirstTableProperties(produced)?.TableWidth;
        Assert.Equal("5000", width?.Width?.Value);
        Assert.Equal(W.TableWidthUnitValues.Dxa, width?.Type?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }
}
