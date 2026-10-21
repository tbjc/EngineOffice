using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// Reusable cell styles: font, solid fill, borders and alignment. Applying one style to several
/// cells reuses the stylesheet components and the <c>cellXfs</c> entry.
/// </summary>
public sealed class ExcelCellStyleTests
{
    [Fact]
    public void Style_NewCell_IsDefaultWithoutCreatingIt()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        CellStyle style = sheet["A1"].Style;

        Assert.Null(style.FontName);
        Assert.Null(style.FontSizePoints);
        Assert.Null(style.Bold);
        Assert.Null(style.Italic);
        Assert.Null(style.Underline);
        Assert.Null(style.FontColor);
        Assert.Null(style.FillColor);
        Assert.Null(style.Borders);
        Assert.Null(style.HorizontalAlignment);
        Assert.Null(style.VerticalAlignment);
        Assert.Null(style.WrapText);

        sheet["A1"].Style = new CellStyle();

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Style_AppliesFontFillBordersAndAlignment_RoundTrips()
    {
        CellStyle expected = new()
        {
            FontName = "Arial",
            FontSizePoints = 12.5,
            Bold = true,
            Italic = true,
            Underline = CellUnderlineStyle.DoubleLine,
            FontColor = RgbColor.FromRgb(0x11, 0x22, 0x33),
            FillColor = RgbColor.FromRgb(0x44, 0x55, 0x66),
            Borders = new CellBorders(
                left: new CellBorder(CellBorderStyle.Thin, RgbColor.FromRgb(0xAA, 0xBB, 0xCC)),
                right: new CellBorder(CellBorderStyle.Medium, RgbColor.FromRgb(0x01, 0x02, 0x03)),
                top: new CellBorder(CellBorderStyle.Dashed),
                bottom: new CellBorder(CellBorderStyle.Thick, RgbColor.FromRgb(0x10, 0x20, 0x30))),
            HorizontalAlignment = CellHorizontalAlignment.Center,
            VerticalAlignment = CellVerticalAlignment.Center,
            WrapText = true,
        };

        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Cell cell = workbook.Worksheets[0]["B2"];
            cell.Value = "styled";
            cell.Style = expected;
            bytes = SaveBytes.ToBytes(workbook);
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell reopenedCell = reopened.Worksheets[0]["B2"];
        CellStyle actual = reopenedCell.Style;

        Assert.Equal("styled", reopenedCell.Value);
        Assert.Equal(expected.FontName, actual.FontName);
        Assert.Equal(expected.FontSizePoints, actual.FontSizePoints);
        Assert.Equal(expected.Bold, actual.Bold);
        Assert.Equal(expected.Italic, actual.Italic);
        Assert.Equal(expected.Underline, actual.Underline);
        Assert.Equal(expected.FontColor, actual.FontColor);
        Assert.Equal(expected.FillColor, actual.FillColor);
        Assert.Equal(expected.Borders, actual.Borders);
        Assert.Equal(expected.HorizontalAlignment, actual.HorizontalAlignment);
        Assert.Equal(expected.VerticalAlignment, actual.VerticalAlignment);
        Assert.Equal(expected.WrapText, actual.WrapText);

        (uint StyleIndex, S.CellFormat Format, S.Font Font, S.Fill Fill, S.Border Border) parts =
            OoxmlValidation.ReadXlsxCellStyleParts(bytes, "Sheet1", "B2");

        Assert.Equal("Arial", parts.Font.FontName?.Val?.Value);
        Assert.Equal(12.5, parts.Font.FontSize?.Val?.Value);
        Assert.True(parts.Font.Bold is not null);
        Assert.True(parts.Font.Italic is not null);
        Assert.Equal(S.UnderlineValues.Double, parts.Font.Underline!.Val!.Value);
        Assert.Equal("FF112233", parts.Font.Color?.Rgb?.Value);
        Assert.Equal("FF445566", parts.Fill.PatternFill?.ForegroundColor?.Rgb?.Value);
        Assert.Equal(S.BorderStyleValues.Thin, parts.Border.LeftBorder!.Style!.Value);
        Assert.Equal("FFAABBCC", parts.Border.LeftBorder!.Color!.Rgb!.Value);
        Assert.Equal(S.BorderStyleValues.Medium, parts.Border.RightBorder!.Style!.Value);
        Assert.Equal(S.BorderStyleValues.Dashed, parts.Border.TopBorder!.Style!.Value);
        Assert.Equal(S.BorderStyleValues.Thick, parts.Border.BottomBorder!.Style!.Value);
        Assert.Equal(S.HorizontalAlignmentValues.Center, parts.Format.Alignment!.Horizontal!.Value);
        Assert.Equal(S.VerticalAlignmentValues.Center, parts.Format.Alignment!.Vertical!.Value);
        Assert.True(parts.Format.Alignment!.WrapText!.Value);
        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Style_SharedAcrossCells_ReusesCellFormatsAndComponents()
    {
        CellStyle style = new()
        {
            FontName = "Arial",
            Bold = true,
            FillColor = RgbColor.FromRgb(0x12, 0x34, 0x56),
            Borders = new CellBorders(bottom: new CellBorder(CellBorderStyle.Thin)),
            HorizontalAlignment = CellHorizontalAlignment.Center,
        };

        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A1"].Style = style;
            sheet["B1"].Style = style;
            sheet["C1"].Style = style;
            bytes = SaveBytes.ToBytes(workbook);
        }

        (int Fonts, int Fills, int Borders, int CellFormats) counts =
            OoxmlValidation.ReadXlsxStyleCounts(bytes);
        Assert.Equal(2, counts.Fonts);
        Assert.Equal(2, counts.Fills);
        Assert.Equal(2, counts.Borders);
        Assert.Equal(2, counts.CellFormats);

        uint first = OoxmlValidation.ReadXlsxCellStyle(bytes, "Sheet1", "A1").StyleIndex;
        uint second = OoxmlValidation.ReadXlsxCellStyle(bytes, "Sheet1", "B1").StyleIndex;
        uint third = OoxmlValidation.ReadXlsxCellStyle(bytes, "Sheet1", "C1").StyleIndex;
        Assert.Equal(first, second);
        Assert.Equal(first, third);
        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Style_DifferentStyles_UseDifferentCellFormats()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A1"].Style = new CellStyle { Bold = true };
            sheet["B1"].Style = new CellStyle { Italic = true };
            bytes = SaveBytes.ToBytes(workbook);
        }

        (int Fonts, int Fills, int Borders, int CellFormats) counts =
            OoxmlValidation.ReadXlsxStyleCounts(bytes);
        Assert.Equal(3, counts.Fonts);
        Assert.Equal(1, counts.Fills);
        Assert.Equal(1, counts.Borders);
        Assert.Equal(3, counts.CellFormats);

        uint first = OoxmlValidation.ReadXlsxCellStyle(bytes, "Sheet1", "A1").StyleIndex;
        uint second = OoxmlValidation.ReadXlsxCellStyle(bytes, "Sheet1", "B1").StyleIndex;
        Assert.NotEqual(first, second);
        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Style_PreservesNumberFormatAndDateValue_RoundTrips()
    {
        RgbColor fill = RgbColor.FromRgb(0xAB, 0xCD, 0xEF);
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.SetDateTime(new DateTime(2024, 1, 31), "yyyy-mm-dd");
            cell.Style = new CellStyle { Bold = true, FillColor = fill };
            bytes = SaveBytes.ToBytes(workbook);
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell reopenedCell = reopened.Worksheets[0]["A1"];
        Assert.Equal(new DateTime(2024, 1, 31), reopenedCell.Value);
        Assert.Equal("yyyy-mm-dd", reopenedCell.NumberFormat);
        Assert.True(reopenedCell.Style.Bold);
        Assert.Equal(fill, reopenedCell.Style.FillColor);
        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Style_DefaultResetsModeledFormattingAndKeepsNumberFormat()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.Value = 1.5;
            cell.NumberFormat = "0.00";
            cell.Style = new CellStyle
            {
                FontName = "Arial",
                Bold = true,
                FillColor = RgbColor.FromRgb(0x11, 0x22, 0x33),
                Borders = new CellBorders(bottom: new CellBorder(CellBorderStyle.Thin)),
                HorizontalAlignment = CellHorizontalAlignment.Center,
                WrapText = true,
            };
            cell.Style = new CellStyle();
            bytes = SaveBytes.ToBytes(workbook);
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell reopenedCell = reopened.Worksheets[0]["A1"];
        CellStyle actual = reopenedCell.Style;

        Assert.Equal(1.5, reopenedCell.Value);
        Assert.Equal("0.00", reopenedCell.NumberFormat);
        Assert.Null(actual.FontName);
        Assert.Null(actual.FontSizePoints);
        Assert.Null(actual.Bold);
        Assert.Null(actual.Italic);
        Assert.Null(actual.Underline);
        Assert.Null(actual.FontColor);
        Assert.Null(actual.FillColor);
        Assert.Null(actual.Borders);
        Assert.Null(actual.HorizontalAlignment);
        Assert.Null(actual.VerticalAlignment);
        Assert.Null(actual.WrapText);
        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Style_NonDefaultOnMissingCell_CreatesCellWithoutValue()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["C3"].Style = new CellStyle { Bold = true };
            bytes = SaveBytes.ToBytes(workbook);
        }

        string sheetXml = TestPackage.ReadEntryText(bytes, "xl/worksheets/sheet1.xml");
        Assert.Contains("C3", sheetXml, StringComparison.Ordinal);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell cell = reopened.Worksheets[0]["C3"];
        Assert.Null(cell.Value);
        Assert.True(cell.Style.Bold);
        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Style_EmptyBordersOnMissingCell_DoesNotCreateCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        sheet["A1"].Style = new CellStyle { Borders = new CellBorders() };

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("bad\u0001name")]
    public void Style_InvalidFontName_ThrowsWithoutCreatingCell(string fontName)
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet["A1"].Style = new CellStyle { FontName = fontName });

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Style_TooLongFontName_ThrowsWithoutCreatingCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(
            () => sheet["A1"].Style = new CellStyle { FontName = new string('x', 256) });

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.0)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(500.0)]
    public void Style_InvalidFontSize_ThrowsWithoutCreatingCell(double fontSize)
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentOutOfRangeException>(
            () => sheet["A1"].Style = new CellStyle { FontSizePoints = fontSize });

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Style_UndefinedEnumValues_ThrowWithoutCreatingCells()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentOutOfRangeException>(
            () => sheet["A1"].Style = new CellStyle { Underline = (CellUnderlineStyle)999 });
        Assert.Throws<ArgumentOutOfRangeException>(
            () => sheet["B1"].Style = new CellStyle { HorizontalAlignment = (CellHorizontalAlignment)999 });
        Assert.Throws<ArgumentOutOfRangeException>(
            () => sheet["C1"].Style = new CellStyle { VerticalAlignment = (CellVerticalAlignment)999 });

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
        Assert.DoesNotContain("B1", sheetXml, StringComparison.Ordinal);
        Assert.DoesNotContain("C1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Style_Null_ThrowsWithoutCreatingCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentNullException>(() => sheet["A1"].Style = null!);

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void CellBorder_UndefinedStyle_Throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new CellBorder((CellBorderStyle)999));
    }

    [Fact]
    public void Style_AllBorderStyles_RoundTrip()
    {
        foreach (CellBorderStyle borderStyle in Enum.GetValues<CellBorderStyle>())
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Create();
            workbook.Worksheets[0]["A1"].Style = new CellStyle
            {
                Borders = new CellBorders(left: new CellBorder(borderStyle)),
            };
            byte[] bytes = SaveBytes.ToBytes(workbook);

            Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));

            using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
            CellBorders? borders = reopened.Worksheets[0]["A1"].Style.Borders;
            if (borderStyle == CellBorderStyle.None)
            {
                Assert.Null(borders);
            }
            else
            {
                Assert.Equal(borderStyle, borders!.Value.Left!.Value.Style);
            }
        }
    }

    [Fact]
    public void Style_AllUnderlineStyles_RoundTrip()
    {
        foreach (CellUnderlineStyle underline in Enum.GetValues<CellUnderlineStyle>())
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Create();
            workbook.Worksheets[0]["A1"].Style = new CellStyle { Underline = underline };
            byte[] bytes = SaveBytes.ToBytes(workbook);

            Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));

            using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
            Assert.Equal(underline, reopened.Worksheets[0]["A1"].Style.Underline);
        }
    }

    [Fact]
    public void Style_AllAlignmentValues_RoundTrip()
    {
        foreach (CellHorizontalAlignment horizontal in Enum.GetValues<CellHorizontalAlignment>())
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Create();
            workbook.Worksheets[0]["A1"].Style = new CellStyle { HorizontalAlignment = horizontal };
            byte[] bytes = SaveBytes.ToBytes(workbook);

            Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));

            using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
            Assert.Equal(horizontal, reopened.Worksheets[0]["A1"].Style.HorizontalAlignment);
        }

        foreach (CellVerticalAlignment vertical in Enum.GetValues<CellVerticalAlignment>())
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Create();
            workbook.Worksheets[0]["A1"].Style = new CellStyle { VerticalAlignment = vertical };
            byte[] bytes = SaveBytes.ToBytes(workbook);

            Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));

            using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
            Assert.Equal(vertical, reopened.Worksheets[0]["A1"].Style.VerticalAlignment);
        }
    }

    [Fact]
    public void Style_ReadsExistingFormattingFromFixture()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell(
            "A1",
            "1",
            0U,
            includeOtherFormatting: true);

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));
        CellStyle style = workbook.Worksheets[0]["A1"].Style;

        Assert.Null(style.FontName);
        Assert.Null(style.FontSizePoints);
        Assert.True(style.Bold);
        Assert.Null(style.Italic);
        Assert.Null(style.Underline);
        Assert.Null(style.FontColor);
        Assert.Equal(RgbColor.FromRgb(0xFF, 0xFF, 0x00), style.FillColor);
        Assert.True(style.Borders.HasValue);

        CellBorders borders = style.Borders.Value;
        Assert.Null(borders.Left);
        Assert.Null(borders.Right);
        Assert.Null(borders.Top);
        Assert.Equal(CellBorderStyle.Thick, borders.Bottom!.Value.Style);
        Assert.Null(borders.Bottom.Value.Color);
        Assert.Equal(CellHorizontalAlignment.Center, style.HorizontalAlignment);
        Assert.Null(style.VerticalAlignment);
        Assert.Null(style.WrapText);
    }

    [Fact]
    public void OpenedWorkbook_WithCellStyle_PreservesItWhenTheValueChanges()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell(
            "A1",
            "1",
            0U,
            includeOtherFormatting: true);

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            workbook.Worksheets[0]["A1"].Value = 2.0;
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
        (uint StyleIndex, S.CellFormat Format, S.Font Font, S.Fill Fill, S.Border Border) parts =
            OoxmlValidation.ReadXlsxCellStyleParts(produced, "Hoja1", "A1");
        Assert.True(parts.Font.Bold is not null);
        Assert.Equal("FFFFFF00", parts.Fill.PatternFill?.ForegroundColor?.Rgb?.Value);
        Assert.Equal(S.BorderStyleValues.Thick, parts.Border.BottomBorder!.Style!.Value);
        Assert.Equal(S.HorizontalAlignmentValues.Center, parts.Format.Alignment!.Horizontal!.Value);
    }

    [Fact]
    public void Style_PreservesUnmodelledFontBorderAndAlignmentProperties()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell(
            "A1",
            "1",
            0U,
            includeOtherFormatting: true,
            includeUnmodelledFormatting: true);

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.Style = new CellStyle
            {
                Italic = true,
                HorizontalAlignment = CellHorizontalAlignment.Right,
            };
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
        (uint StyleIndex, S.CellFormat Format, S.Font Font, S.Fill Fill, S.Border Border) parts =
            OoxmlValidation.ReadXlsxCellStyleParts(produced, "Hoja1", "A1");

        Assert.True(parts.Font.Italic is not null);
        Assert.True(parts.Font.Strike is not null);
        Assert.Null(parts.Font.Bold);
        Assert.NotNull(parts.Border.DiagonalBorder);
        Assert.Null(parts.Border.BottomBorder);
        Assert.Equal(S.HorizontalAlignmentValues.Right, parts.Format.Alignment!.Horizontal!.Value);
        Assert.Equal(2U, parts.Format.Alignment!.Indent!.Value);
    }

    [Fact]
    public void DisposedWorkbook_StyleThrowsObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        Cell cell = workbook.Worksheets[0]["A1"];
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = cell.Style);
        Assert.Throws<ObjectDisposedException>(() => cell.Style = new CellStyle());
    }
}
