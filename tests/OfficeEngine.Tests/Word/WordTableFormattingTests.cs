using System;
using System.Linq;
using OfficeEngine;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.5b: table and cell formatting — preferred width, alignment, layout, borders, shading
/// and vertical alignment. Each property reads, writes and clears its own markup, and values Office
/// Open XML defines but this version does not model are reported instead of guessed.
/// </summary>
public sealed class WordTableFormattingTests
{
    [Fact]
    public void NewTable_HasNoFormatting()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);
        TableCell cell = table.Rows[0].Cells[0];

        Assert.Null(table.Width);
        Assert.Null(table.Alignment);
        Assert.Null(table.Layout);
        Assert.Null(table.ShadingColor);
        Assert.Null(table.BorderTop);

        Assert.Null(cell.Width);
        Assert.Null(cell.VerticalAlignment);
        Assert.Null(cell.ShadingColor);
        Assert.Null(cell.BorderTop);

        // Formatting is opt-in: the created table keeps an empty w:tblPr and no w:tcPr at all.
        W.TableProperties properties = Assert.IsType<W.TableProperties>(
            OoxmlValidation.ReadDocxFirstTableProperties(SaveBytes.ToBytes(document)));
        Assert.Empty(properties.ChildElements);
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
    }

    [Theory]
    [InlineData(120.0, "2400")]
    [InlineData(0.05, "1")]
    public void Table_WidthPoints_IsStoredInTwips(double points, string twips)
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);

        table.Width = TableWidth.Points(points);

        Assert.Equal(TableWidthType.Points, table.Width!.Value.Type);
        Assert.Equal(points, table.Width.Value.Value);

        byte[] produced = SaveBytes.ToBytes(document);
        W.TableWidth? stored = OoxmlValidation.ReadDocxFirstTableProperties(produced)?.TableWidth;
        Assert.Equal(twips, stored?.Width?.Value);
        Assert.Equal(W.TableWidthUnitValues.Dxa, stored?.Type?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Table_Width_SupportsPercentageAndAuto()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);

        table.Width = TableWidth.Percentage(25.0);
        Assert.Equal(TableWidthType.Percentage, table.Width!.Value.Type);
        Assert.Equal(25.0, table.Width.Value.Value);

        table.Width = TableWidth.Auto;
        Assert.Equal(TableWidthType.Auto, table.Width!.Value.Type);

        table.Width = null;
        Assert.Null(table.Width);
    }

    [Fact]
    public void Table_Width_WithoutType_IsReadAsTwips()
    {
        // ST_TblWidth: "If this attribute is omitted, then its value shall be assumed to be dxa."
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableWidth { Width = "2400" }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(TableWidth.Points(120.0), document.Tables[0].Width);
    }

    [Fact]
    public void Table_Width_WithoutMeasure_ReadsAsZero()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableWidth { Type = W.TableWidthUnitValues.Dxa }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(TableWidth.Points(0.0), document.Tables[0].Width);
    }

    [Fact]
    public void Table_Width_ReportsAnUnmodelledUnit()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            OoxmlValidation.WithAttribute(new W.TableWidth { Width = "2400" }, "type", "nil")));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<UnsupportedOfficeFeatureException>(() => document.Tables[0].Width);
    }

    [Fact]
    public void Table_Width_ReportsAMalformedMeasure()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableWidth { Width = "wide", Type = W.TableWidthUnitValues.Dxa }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Tables[0].Width);
    }

    [Fact]
    public void TableWidth_RejectsInvalidPoints()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => TableWidth.Points(-1.0));
        Assert.Throws<ArgumentOutOfRangeException>(() => TableWidth.Points(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => TableWidth.Percentage(-1.0));
        Assert.Throws<ArgumentOutOfRangeException>(() => TableWidth.Percentage(double.PositiveInfinity));
    }

    [Fact]
    public void Table_Alignment_CanBeSetReadAndCleared()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);

        foreach (TableAlignment alignment in Enum.GetValues<TableAlignment>())
        {
            table.Alignment = alignment;
            Assert.Equal(alignment, table.Alignment);
        }

        table.Alignment = null;
        Assert.Null(table.Alignment);
    }

    [Fact]
    public void Table_Alignment_IsStoredAsTheJustificationElement()
    {
        using WordDocument document = WordDocument.Create();
        document.AddTable(1, 1).Alignment = TableAlignment.Center;

        byte[] produced = SaveBytes.ToBytes(document);

        Assert.Equal(
            W.TableRowAlignmentValues.Center,
            OoxmlValidation.ReadDocxFirstTableProperties(produced)?.TableJustification?.Val?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Table_Alignment_ReportsAnUnmodelledValue()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            OoxmlValidation.WithAttribute(new W.TableJustification(), "val", "start")));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Tables[0].Alignment);
    }

    [Fact]
    public void Table_Layout_CanBeSetReadAndCleared()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);

        table.Layout = TableLayout.Fixed;
        Assert.Equal(TableLayout.Fixed, table.Layout);

        table.Layout = TableLayout.AutoFit;
        Assert.Equal(TableLayout.AutoFit, table.Layout);

        table.Layout = null;
        Assert.Null(table.Layout);
    }

    [Fact]
    public void Table_Layout_WithoutType_IsReadAsAutoFit()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableLayout()));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(TableLayout.AutoFit, document.Tables[0].Layout);
    }

    [Fact]
    public void Table_Layout_ReportsAnUnmodelledValue()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            OoxmlValidation.WithAttribute(new W.TableLayout(), "type", "distribute")));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Tables[0].Layout);
    }

    [Fact]
    public void Table_AlignmentAndLayout_RejectUndefinedValues()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);

        Assert.Throws<ArgumentOutOfRangeException>(() => table.Alignment = (TableAlignment)99);
        Assert.Throws<ArgumentOutOfRangeException>(() => table.Layout = (TableLayout)99);

        // A rejected value must leave no partial markup behind.
        Assert.Null(table.Alignment);
        Assert.Null(table.Layout);
    }

    [Fact]
    public void Table_ShadingColor_IsAStoredSolidFill()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);

        table.ShadingColor = new RgbColor(0xD9, 0xE2, 0xF3);

        Assert.Equal(new RgbColor(0xD9, 0xE2, 0xF3), table.ShadingColor);

        byte[] produced = SaveBytes.ToBytes(document);
        W.Shading? shading = OoxmlValidation.ReadDocxFirstTableProperties(produced)?.Shading;
        Assert.Equal(W.ShadingPatternValues.Clear, shading?.Val?.Value);
        Assert.Equal("D9E2F3", shading?.Fill?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        table.ShadingColor = null;
        Assert.Null(table.ShadingColor);
    }

    [Fact]
    public void Table_ShadingColor_IsNullForAnAutomaticOrThemedFill()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "auto" }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Null(document.Tables[0].ShadingColor);
    }

    [Fact]
    public void Table_ShadingColor_ReportsAnUnmodelledPattern()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.Shading { Val = W.ShadingPatternValues.Percent10, Fill = "FF0000" }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<UnsupportedOfficeFeatureException>(() => document.Tables[0].ShadingColor);
    }

    [Fact]
    public void Table_ShadingColor_ReportsMissingAndMalformedValues()
    {
        byte[] missingPattern = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.Shading { Fill = "FF0000" }));
        byte[] malformedFill = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "not-a-colour" }));

        using WordDocument withoutPattern = WordDocument.Open(TestPackage.ToStream(missingPattern));
        using WordDocument withBadFill = WordDocument.Open(TestPackage.ToStream(malformedFill));

        Assert.Throws<InvalidOfficeFileException>(() => withoutPattern.Tables[0].ShadingColor);
        Assert.Throws<InvalidOfficeFileException>(() => withBadFill.Tables[0].ShadingColor);
    }

    [Fact]
    public void Table_Borders_CanBeSetReadAndClearedPerEdge()
    {
        using WordDocument document = WordDocument.Create();
        Table table = document.AddTable(1, 1);
        TableBorder border = new(BorderStyle.SingleLine, 1.5, new RgbColor(0xFF, 0x00, 0x00));

        table.BorderTop = border;
        table.BorderLeft = border;
        table.BorderBottom = border;
        table.BorderRight = border;
        table.BorderInsideHorizontal = border;
        table.BorderInsideVertical = border;

        Assert.Equal(border, table.BorderTop);
        Assert.Equal(border, table.BorderLeft);
        Assert.Equal(border, table.BorderBottom);
        Assert.Equal(border, table.BorderRight);
        Assert.Equal(border, table.BorderInsideHorizontal);
        Assert.Equal(border, table.BorderInsideVertical);

        table.BorderTop = null;
        Assert.Null(table.BorderTop);
        Assert.Equal(border, table.BorderBottom);

        table.BorderLeft = null;
        table.BorderBottom = null;
        table.BorderRight = null;
        table.BorderInsideHorizontal = null;
        table.BorderInsideVertical = null;

        // Clearing the last edge removes the container as well.
        Assert.Null(OoxmlValidation.ReadDocxFirstTableProperties(SaveBytes.ToBytes(document))?.TableBorders);
    }

    [Fact]
    public void Table_Border_IsWrittenInEighthsOfAPoint()
    {
        using WordDocument document = WordDocument.Create();
        document.AddTable(1, 1).BorderTop = new TableBorder(BorderStyle.DoubleLine, 1.5);

        byte[] produced = SaveBytes.ToBytes(document);
        W.TopBorder? top = OoxmlValidation.ReadDocxFirstTableProperties(produced)?.TableBorders?.TopBorder;

        Assert.Equal(W.BorderValues.Double, top?.Val?.Value);
        Assert.Equal(12U, top?.Size?.Value);
        Assert.Equal("auto", top?.Color?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void Table_Border_EveryStyleRoundTripsThroughTheDocument()
    {
        foreach (BorderStyle style in Enum.GetValues<BorderStyle>())
        {
            using WordDocument document = WordDocument.Create();
            Table table = document.AddTable(1, 1);
            table.BorderTop = new TableBorder(style, 0.5);

            Assert.Equal(style, table.BorderTop!.Value.Style);
        }
    }

    [Fact]
    public void Table_Border_DefaultsToHalfAPointWhenTheSizeIsAbsent()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableBorders(new W.TopBorder { Val = W.BorderValues.Single })));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Equal(new TableBorder(BorderStyle.SingleLine, 0.5), document.Tables[0].BorderTop);
    }

    [Fact]
    public void Table_Border_ReportsUnmodelledOrMalformedMarkup()
    {
        byte[] artStyle = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableBorders(OoxmlValidation.WithAttribute(new W.TopBorder(), "val", "apples"))));
        byte[] missingStyle = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableBorders(new W.TopBorder())));
        byte[] malformedSize = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableBorders(OoxmlValidation.WithAttribute(
                new W.TopBorder { Val = W.BorderValues.Single }, "sz", "thick"))));

        using WordDocument withArt = WordDocument.Open(TestPackage.ToStream(artStyle));
        using WordDocument withoutStyle = WordDocument.Open(TestPackage.ToStream(missingStyle));
        using WordDocument withBadSize = WordDocument.Open(TestPackage.ToStream(malformedSize));

        Assert.Throws<UnsupportedOfficeFeatureException>(() => withArt.Tables[0].BorderTop);
        Assert.Throws<InvalidOfficeFileException>(() => withoutStyle.Tables[0].BorderTop);
        Assert.Throws<InvalidOfficeFileException>(() => withBadSize.Tables[0].BorderTop);
    }

    [Fact]
    public void TableBorder_RejectsInvalidArguments()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new TableBorder((BorderStyle)99, 1.0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new TableBorder(BorderStyle.SingleLine, -1.0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new TableBorder(BorderStyle.SingleLine, double.NaN));
    }

    [Fact]
    public void TableBorder_ReportsItsValuesAsText()
    {
        TableBorder plain = new(BorderStyle.Dashed, 0.5);
        TableBorder coloured = new(BorderStyle.Dashed, 0.5, RgbColor.TryParseHex("#336699", out RgbColor blue) ? blue : default);

        Assert.Equal("Dashed 0.5 pt", plain.ToString());
        Assert.Equal("Dashed 0.5 pt #336699", coloured.ToString());
        Assert.Equal(plain, new TableBorder(BorderStyle.Dashed, 0.5));
        Assert.NotEqual(plain, coloured);
    }

    [Fact]
    public void Table_Formatting_RoundTripsThroughSaveAndOpen()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            Table table = document.AddTable(1, 1);
            table.Width = TableWidth.Points(300.0);
            table.Alignment = TableAlignment.Right;
            table.Layout = TableLayout.Fixed;
            table.ShadingColor = new RgbColor(0x00, 0xFF, 0x00);
            table.BorderBottom = new TableBorder(BorderStyle.Dotted, 0.5);
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Table read = reopened.Tables[0];

        Assert.Equal(TableWidth.Points(300.0), read.Width);
        Assert.Equal(TableAlignment.Right, read.Alignment);
        Assert.Equal(TableLayout.Fixed, read.Layout);
        Assert.Equal(new RgbColor(0x00, 0xFF, 0x00), read.ShadingColor);
        Assert.Equal(new TableBorder(BorderStyle.Dotted, 0.5), read.BorderBottom);
    }

    [Fact]
    public void Table_Formatting_IsInsertedInSchemaOrderAroundUnmodelledMarkup()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableIndentation { Width = 120 },
            new W.TableLook { Val = "04A0" }));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Table table = document.Tables[0];
        table.Width = TableWidth.Points(200.0);
        table.Alignment = TableAlignment.Center;
        table.BorderTop = new TableBorder(BorderStyle.SingleLine, 0.5);
        table.ShadingColor = new RgbColor(0xFF, 0xFF, 0x00);
        table.Layout = TableLayout.Fixed;

        W.TableProperties properties = Assert.IsType<W.TableProperties>(
            OoxmlValidation.ReadDocxFirstTableProperties(SaveBytes.ToBytes(document)));

        Assert.Equal(
            ["tblW", "jc", "tblInd", "tblBorders", "shd", "tblLayout", "tblLook"],
            properties.ChildElements.Select(child => child.LocalName));
    }

    [Fact]
    public void TableCell_WidthAndVerticalAlignment_CanBeSetReadAndCleared()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        cell.Width = TableWidth.Points(60.0);
        cell.VerticalAlignment = TableVerticalAlignment.Center;

        Assert.Equal(TableWidth.Points(60.0), cell.Width);
        Assert.Equal(TableVerticalAlignment.Center, cell.VerticalAlignment);

        byte[] produced = SaveBytes.ToBytes(document);
        W.TableCellWidth? width = OoxmlValidation.ReadDocxFirstTableCellProperties(produced)?.TableCellWidth;
        Assert.Equal(1200, width?.Width?.Value is string twips
            ? int.Parse(twips, System.Globalization.CultureInfo.InvariantCulture)
            : 0);
        Assert.Equal(W.TableWidthUnitValues.Dxa, width?.Type?.Value);
        Assert.Equal(
            W.TableVerticalAlignmentValues.Center,
            OoxmlValidation.ReadDocxFirstTableCellProperties(produced)?.TableCellVerticalAlignment?.Val?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void TableCell_ClearingTheLastProperty_RemovesTheEmptyPropertiesElement()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        cell.Width = TableWidth.Points(60.0);
        cell.ShadingColor = new RgbColor(0xAB, 0xC1, 0x23);
        cell.BorderTop = new TableBorder(BorderStyle.SingleLine, 0.5);

        cell.Width = null;
        cell.ShadingColor = null;
        cell.BorderTop = null;

        Assert.Null(cell.Width);
        Assert.Null(cell.ShadingColor);
        Assert.Null(cell.BorderTop);
        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
        Assert.Empty(OoxmlValidation.ValidateDocx(SaveBytes.ToBytes(document)));
    }

    [Fact]
    public void TableCell_VerticalAlignment_ReportsUnmodelledOrMalformedMarkup()
    {
        byte[] unmodelled = OoxmlValidation.CreateDocxWithCellProperties("cell", new W.TableCellProperties(
            OoxmlValidation.WithAttribute(new W.TableCellVerticalAlignment(), "val", "baseline")));
        byte[] missingVal = OoxmlValidation.CreateDocxWithCellProperties("cell", new W.TableCellProperties(
            new W.TableCellVerticalAlignment()));

        using WordDocument withUnmodelled = WordDocument.Open(TestPackage.ToStream(unmodelled));
        using WordDocument withoutVal = WordDocument.Open(TestPackage.ToStream(missingVal));

        Assert.Throws<InvalidOfficeFileException>(() => withUnmodelled.Tables[0].Rows[0].Cells[0].VerticalAlignment);
        Assert.Null(withoutVal.Tables[0].Rows[0].Cells[0].VerticalAlignment);
    }

    [Fact]
    public void TableCell_VerticalAlignment_RejectsUndefinedValues()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        Assert.Throws<ArgumentOutOfRangeException>(() => cell.VerticalAlignment = (TableVerticalAlignment)99);
        Assert.Null(cell.VerticalAlignment);
    }

    [Fact]
    public void TableCell_Borders_NeverTouchTheEdgesItDoesNotOwn()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithCellProperties("cell", new W.TableCellProperties(
            new W.TableCellBorders(
                new W.InsideHorizontalBorder { Val = W.BorderValues.Dotted },
                new W.TopLeftToBottomRightCellBorder { Val = W.BorderValues.Dashed })));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        TableCell cell = document.Tables[0].Rows[0].Cells[0];

        cell.BorderTop = new TableBorder(BorderStyle.SingleLine, 0.5);
        cell.BorderTop = null;

        W.TableCellBorders borders = Assert.IsType<W.TableCellBorders>(
            OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document))?.TableCellBorders);

        Assert.Equal(
            ["insideH", "tl2br"],
            borders.ChildElements.Select(child => child.LocalName));
    }

    [Fact]
    public void TableWidth_ReportsItsMeasureAndRejectsOverflow()
    {
        Assert.Equal(TableWidthType.Points, TableWidth.Points(120.0).Type);
        Assert.Equal(120.0, TableWidth.Points(120.0).Value);
        Assert.Equal(TableWidthType.Percentage, TableWidth.Percentage(25.5).Type);
        Assert.Equal(25.5, TableWidth.Percentage(25.5).Value);
        Assert.Equal(TableWidthType.Auto, TableWidth.Auto.Type);
        Assert.Equal(0.0, TableWidth.Auto.Value);

        Assert.Equal("120 pt", TableWidth.Points(120.0).ToString());
        Assert.Equal("25.5 %", TableWidth.Percentage(25.5).ToString());
        Assert.Equal("auto", TableWidth.Auto.ToString());

        Assert.True(TableWidth.Points(120.0) == TableWidth.Points(120.0));
        Assert.True(TableWidth.Points(120.0) != TableWidth.Auto);
        Assert.Equal(TableWidth.Points(120.0).GetHashCode(), TableWidth.Points(120.0).GetHashCode());

        Assert.Throws<ArgumentOutOfRangeException>(() => TableWidth.Points(double.MaxValue));
        Assert.Throws<ArgumentOutOfRangeException>(() => TableWidth.Percentage(double.MaxValue));
    }

    [Fact]
    public void Table_Width_RoundTripsPercentagesAndReportsUnreadableMeasures()
    {
        using MemoryStream stream = new();
        using (WordDocument document = WordDocument.Create())
        {
            document.AddTable(1, 1).Width = TableWidth.Percentage(33.4);
            document.Save(stream);
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Open(stream);
        Assert.Equal(TableWidth.Percentage(33.4), reopened.Tables[0].Width);

        byte[] unknownUnit = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            OoxmlValidation.WithAttribute(new W.TableWidth { Width = "2400" }, "type", "bogus")));
        byte[] negative = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableWidth { Width = "-2400", Type = W.TableWidthUnitValues.Dxa }));

        using WordDocument withUnknownUnit = WordDocument.Open(TestPackage.ToStream(unknownUnit));
        using WordDocument withNegative = WordDocument.Open(TestPackage.ToStream(negative));

        Assert.Throws<InvalidOfficeFileException>(() => withUnknownUnit.Tables[0].Width);
        Assert.Throws<InvalidOfficeFileException>(() => withNegative.Tables[0].Width);
    }

    [Fact]
    public void Table_WithoutProperties_ReadsNothingAndGainsThemOnWrite()
    {
        // A malformed package can lack w:tblPr; reading must not fail and writing must repair the order.
        byte[] fixture = OoxmlValidation.CreateDocxWithBody(new W.Table(
            new W.TableGrid(new W.GridColumn { Width = "2400" }),
            new W.TableRow(new W.TableCell(new W.Paragraph(new W.Run(new W.Text("cell")))))));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        Table table = document.Tables[0];

        Assert.Null(table.Width);
        Assert.Null(table.Alignment);
        Assert.Null(table.BorderTop);
        Assert.Null(table.ShadingColor);

        // Clearing a property that is not there is a no-op, not a failure.
        table.BorderTop = null;
        table.Width = null;

        table.Width = TableWidth.Points(100.0);

        byte[] produced = SaveBytes.ToBytes(document);
        W.Table read = Assert.Single(OoxmlValidation.ReadDocxTables(produced));
        Assert.Equal("tblPr", read.ChildElements[0].LocalName);
        Assert.Equal("2000", read.TableProperties?.TableWidth?.Width?.Value);
        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
    }

    [Fact]
    public void TableBorder_RejectsASizeTooLargeForTheSchema()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new TableBorder(BorderStyle.SingleLine, 1e12));
    }

    [Fact]
    public void TableBorder_SupportsEqualityOperatorsAndHashing()
    {
        TableBorder plain = new(BorderStyle.Dotted, 0.5);
        TableBorder same = new(BorderStyle.Dotted, 0.5);

        Assert.True(plain == same);
        Assert.False(plain != same);
        Assert.Equal(plain.GetHashCode(), same.GetHashCode());
    }

    [Fact]
    public void Table_Border_ReportsAMalformedColour()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithTableProperties("cell", new W.TableProperties(
            new W.TableBorders(OoxmlValidation.WithAttribute(
                new W.TopBorder { Val = W.BorderValues.Single }, "color", "not-a-colour"))));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Tables[0].BorderTop);
    }

    [Fact]
    public void TableCell_Borders_AllFourOuterEdgesRoundTrip()
    {
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        TableBorder border = new(BorderStyle.Thick, 1.0);

        cell.BorderTop = border;
        cell.BorderLeft = border;
        cell.BorderBottom = border;
        cell.BorderRight = border;

        Assert.Equal(border, cell.BorderTop);
        Assert.Equal(border, cell.BorderLeft);
        Assert.Equal(border, cell.BorderBottom);
        Assert.Equal(border, cell.BorderRight);

        cell.BorderTop = null;
        Assert.Null(cell.BorderTop);
        Assert.Equal(border, cell.BorderLeft);
    }

    [Fact]
    public void TableCell_VerticalAlignment_EveryValueRoundTripsThroughTheDocument()
    {
        foreach (TableVerticalAlignment alignment in Enum.GetValues<TableVerticalAlignment>())
        {
            using WordDocument document = WordDocument.Create();
            TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
            cell.VerticalAlignment = alignment;

            byte[] produced = SaveBytes.ToBytes(document);
            using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));

            Assert.Equal(alignment, reopened.Tables[0].Rows[0].Cells[0].VerticalAlignment);
            Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        }
    }

    [Fact]
    public void TableCell_Formatting_IsInsertedInSchemaOrderAroundUnmodelledMarkup()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithCellProperties("cell", new W.TableCellProperties(
            new W.VerticalMerge { Val = W.MergedCellValues.Continue },
            new W.TableCellMargin()));

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));
        TableCell cell = document.Tables[0].Rows[0].Cells[0];
        cell.Width = TableWidth.Points(60.0);
        cell.BorderTop = new TableBorder(BorderStyle.SingleLine, 0.5);
        cell.ShadingColor = new RgbColor(0xAB, 0xCD, 0xEF);
        cell.VerticalAlignment = TableVerticalAlignment.Bottom;

        W.TableCellProperties properties = Assert.IsType<W.TableCellProperties>(
            OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));

        Assert.Equal(
            ["tcW", "vMerge", "tcBorders", "shd", "tcMar", "vAlign"],
            properties.ChildElements.Select(child => child.LocalName));
    }

    [Fact]
    public void TableCell_ClearingWhenThereIsNothingToClear_IsANoOp()
    {
        // A cell created by AddTable has no w:tcPr, so clearing must not fail nor create one.
        using WordDocument document = WordDocument.Create();
        TableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        cell.Width = null;
        cell.VerticalAlignment = null;
        cell.ShadingColor = null;
        cell.BorderTop = null;
        cell.BorderLeft = null;
        cell.BorderBottom = null;
        cell.BorderRight = null;

        Assert.Null(OoxmlValidation.ReadDocxFirstTableCellProperties(SaveBytes.ToBytes(document)));
    }
}
