using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// Date/time serials and number formats. A date needs an explicit date/time format; reading a
/// numeric cell with such a format returns a DateTime, and styles this version does not model are
/// preserved.
/// </summary>
public sealed class ExcelDateAndNumberFormatTests
{
    [Fact]
    public void NumberFormat_NewCell_IsGeneralWithoutCreatingIt()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Equal("General", sheet["A1"].NumberFormat);

        sheet["A1"].NumberFormat = null;

        Assert.Equal("General", sheet["A1"].NumberFormat);
        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void SetDateTime_WritesSerialAndFormat_RoundTrips()
    {
        DateTime expected = new(2024, 1, 31, 0, 0, 0, DateTimeKind.Utc);
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["B2"].SetDateTime(expected, "yyyy-mm-dd");
            bytes = SaveBytes.ToBytes(workbook);
        }

        string stylesXml = TestPackage.ReadEntryText(bytes, "xl/styles.xml");
        Assert.Equal("45322", OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "B2"));
        Assert.Contains("formatCode=\"yyyy-mm-dd\"", stylesXml, StringComparison.Ordinal);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell cell = reopened.Worksheets[0]["B2"];
        DateTime actual = Assert.IsType<DateTime>(cell.Value);

        Assert.Equal(expected.Ticks, actual.Ticks);
        Assert.Equal(DateTimeKind.Unspecified, actual.Kind);
        Assert.Equal("yyyy-mm-dd", cell.NumberFormat);
    }

    [Fact]
    public void SetDateTime_WithTime_RoundTripsWithinOneMillisecond()
    {
        DateTime expected = new(2024, 2, 29, 13, 45, 30, 250);
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["A1"].SetDateTime(expected, "yyyy-mm-dd hh:mm:ss.000");
            bytes = SaveBytes.ToBytes(workbook);
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell cell = reopened.Worksheets[0]["A1"];
        DateTime actual = Assert.IsType<DateTime>(cell.Value);

        Assert.InRange(actual, expected.AddMilliseconds(-1), expected.AddMilliseconds(1));
        Assert.Equal("yyyy-mm-dd hh:mm:ss.000", cell.NumberFormat);
    }

    [Theory]
    [InlineData(1900, 1, 1, "1")]
    [InlineData(1900, 2, 28, "59")]
    [InlineData(1900, 3, 1, "61")]
    public void SetDateTime_UsesThe1900DateSystemWithItsLeapYearQuirk(
        int year,
        int month,
        int day,
        string expectedSerial)
    {
        DateTime expected = new(year, month, day);
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            workbook.Worksheets[0]["A1"].SetDateTime(expected, "yyyy-mm-dd");
            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Equal(expectedSerial, OoxmlValidation.ReadXlsxCellText(bytes, "Sheet1", "A1"));

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Assert.Equal(expected, Assert.IsType<DateTime>(reopened.Worksheets[0]["A1"].Value));
    }

    [Theory]
    [InlineData("0.00")]
    [InlineData("General")]
    [InlineData("[Red]0.00")]
    [InlineData("\"date\"")]
    public void SetDateTime_RejectsNonDateOrTimeFormatWithoutCreatingTheCell(string format)
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet["A1"].SetDateTime(new DateTime(2024, 1, 1), format));

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void SetDateTime_RejectsNullAndEmptyFormat()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentNullException>(() => sheet["A1"].SetDateTime(new DateTime(2024, 1, 1), null!));
        Assert.Throws<ArgumentException>(() => sheet["A1"].SetDateTime(new DateTime(2024, 1, 1), string.Empty));
    }

    [Fact]
    public void SetDateTime_RejectsDatesBeforeTheWorkbookEpochWithoutCreatingTheCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentOutOfRangeException>(
            () => sheet["A1"].SetDateTime(new DateTime(1899, 12, 31), "yyyy-mm-dd"));

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Value_AssignDateTimeWithoutDateFormat_ThrowsAndDoesNotCreateTheCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<InvalidOperationException>(
            () => sheet["A1"].Value = new DateTime(2024, 1, 1));

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void Value_AssignDateTimeWithNonDateFormat_ThrowsAndKeepsTheValue()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet["A1"].NumberFormat = "0.00";

        Assert.Throws<InvalidOperationException>(
            () => sheet["A1"].Value = new DateTime(2024, 1, 1));

        Assert.Null(sheet["A1"].Value);
        Assert.Equal("0.00", sheet["A1"].NumberFormat);
    }

    [Fact]
    public void Value_AssignDateTimeAfterSettingDateFormat_WritesSerial()
    {
        DateTime expected = new(2024, 3, 15, 8, 30, 0);
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.NumberFormat = "yyyy-mm-dd hh:mm";
            cell.Value = expected;
            bytes = SaveBytes.ToBytes(workbook);
        }

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell reopenedCell = reopened.Worksheets[0]["A1"];
        Assert.Equal(expected, Assert.IsType<DateTime>(reopenedCell.Value));
        Assert.Equal("yyyy-mm-dd hh:mm", reopenedCell.NumberFormat);
    }

    [Fact]
    public void NumberFormat_WritesNumberFormatWithoutChangingTheValue()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.Value = 1.5;
            cell.NumberFormat = "0.00";
            bytes = SaveBytes.ToBytes(workbook);
        }

        string stylesXml = TestPackage.ReadEntryText(bytes, "xl/styles.xml");
        Assert.Contains("formatCode=\"0.00\"", stylesXml, StringComparison.Ordinal);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Cell reopenedCell = reopened.Worksheets[0]["A1"];
        Assert.Equal(1.5, Assert.IsType<double>(reopenedCell.Value));
        Assert.Equal("0.00", reopenedCell.NumberFormat);
    }

    [Fact]
    public void NumberFormat_NullOrGeneral_ResetsToGeneralAndKeepsTheValue()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Cell cell = workbook.Worksheets[0]["A1"];
        cell.Value = 1.5;
        cell.NumberFormat = "0.00";

        cell.NumberFormat = null;
        cell.NumberFormat = null;

        Assert.Equal("General", cell.NumberFormat);
        Assert.Equal(1.5, Assert.IsType<double>(cell.Value));

        cell.NumberFormat = "0.00";
        cell.NumberFormat = "General";

        Assert.Equal("General", cell.NumberFormat);
        Assert.Equal(1.5, Assert.IsType<double>(cell.Value));
    }

    [Fact]
    public void NumberFormat_GeneralOnExistingUnformattedCell_IsANoOp()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Cell cell = workbook.Worksheets[0]["A1"];
        cell.Value = 1.5;

        cell.NumberFormat = null;

        Assert.Equal("General", cell.NumberFormat);
        Assert.Equal(1.5, Assert.IsType<double>(cell.Value));
    }

    [Fact]
    public void NumberFormat_SettingTheSameFormatTwice_IsANoOp()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Cell cell = workbook.Worksheets[0]["A1"];
        cell.NumberFormat = "0.00";

        cell.NumberFormat = "0.00";

        Assert.Equal("0.00", cell.NumberFormat);
    }

    [Fact]
    public void NumberFormat_OnMissingCell_CreatesAFormattedEmptyCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        sheet["A1"].NumberFormat = "0.00";

        Assert.Null(sheet["A1"].Value);
        Assert.Equal("0.00", sheet["A1"].NumberFormat);
        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.Contains("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void NumberFormat_InvalidFormat_ThrowsWithoutCreatingTheCell()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet["A1"].NumberFormat = string.Empty);
        Assert.Throws<ArgumentException>(() => sheet["A1"].NumberFormat = new string('0', 256));
        Assert.Throws<ArgumentException>(() => sheet["A1"].NumberFormat = "\u0001");

        string sheetXml = TestPackage.ReadEntryText(SaveBytes.ToBytes(workbook), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("A1", sheetXml, StringComparison.Ordinal);
    }

    [Fact]
    public void NumberFormat_IsPreservedWhenTheValueChanges()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Cell cell = workbook.Worksheets[0]["A1"];
        cell.NumberFormat = "0.00";

        cell.Value = "text";
        Assert.Equal("0.00", cell.NumberFormat);

        cell.Value = 2.5;
        Assert.Equal("0.00", cell.NumberFormat);
        Assert.Equal(2.5, Assert.IsType<double>(cell.Value));
    }

    [Fact]
    public void OpenedWorkbook_WithBuiltInDateFormat_ReadsDateTimeAndFormat()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell("A1", "45292", 14);
        Assert.Empty(OoxmlValidation.ValidateXlsx(fixture));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));
        Cell cell = workbook.Worksheets[0]["A1"];

        Assert.Equal(new DateTime(2024, 1, 1), Assert.IsType<DateTime>(cell.Value));
        Assert.Equal("mm-dd-yy", cell.NumberFormat);
    }

    [Theory]
    [InlineData("yyyy-mm-dd")]
    [InlineData("mm:ss")]
    [InlineData("[h]:mm:ss")]
    [InlineData("[$-409]m/d/yyyy")]
    [InlineData("h:mm AM/PM")]
    [InlineData("h:mm A/P")]
    [InlineData("A/P")]
    [InlineData("e")]
    public void OpenedWorkbook_WithCustomDateFormat_ReadsDateTime(string format)
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell("A1", "45292", 164, format);
        Assert.Empty(OoxmlValidation.ValidateXlsx(fixture));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));
        Cell cell = workbook.Worksheets[0]["A1"];

        Assert.Equal(new DateTime(2024, 1, 1), Assert.IsType<DateTime>(cell.Value));
        Assert.Equal(format, cell.NumberFormat);
    }

    [Theory]
    [InlineData("0.00E+00")]
    [InlineData("General")]
    [InlineData("[Red]0.00")]
    [InlineData("\"d\"")]
    [InlineData("\\d")]
    [InlineData("0.00 \"days\"")]
    [InlineData("[$-409]General")]
    [InlineData("\"unterminated")]
    [InlineData("[unterminated")]
    [InlineData("[]0.00")]
    [InlineData("\"a\"\"b\"")]
    public void OpenedWorkbook_WithNonDateCustomFormat_ReturnsTheNumber(string format)
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell("A1", "1250", 164, format);
        Assert.Empty(OoxmlValidation.ValidateXlsx(fixture));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));
        Cell cell = workbook.Worksheets[0]["A1"];

        Assert.Equal(1250.0, Assert.IsType<double>(cell.Value));
        Assert.Equal(format, cell.NumberFormat);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("0.5")]
    [InlineData("60")]
    [InlineData("60.5")]
    public void OpenedWorkbook_WithUnrepresentableDateSerial_ReturnsTheSerial(string serial)
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell("A1", serial, 164, "yyyy-mm-dd");

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));
        Cell cell = workbook.Worksheets[0]["A1"];

        Assert.Equal(double.Parse(serial, System.Globalization.CultureInfo.InvariantCulture),
            Assert.IsType<double>(cell.Value));
        Assert.Equal("yyyy-mm-dd", cell.NumberFormat);
    }

    [Fact]
    public void OpenedWorkbook_With1904DateSystem_RoundTripsDatesInThatEpoch()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell(
            "A1", "43830", 164, "yyyy-mm-dd", date1904: true);
        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Worksheet sheet = workbook.Worksheets[0];
            Assert.Equal(new DateTime(2024, 1, 1), Assert.IsType<DateTime>(sheet["A1"].Value));
            Assert.Throws<ArgumentOutOfRangeException>(
                () => sheet["B1"].SetDateTime(new DateTime(1903, 12, 31), "yyyy-mm-dd"));

            sheet["B1"].SetDateTime(new DateTime(2024, 1, 2), "yyyy-mm-dd");
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Equal("43831", OoxmlValidation.ReadXlsxCellText(produced, "Hoja1", "B1"));

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(produced));
        Assert.Equal(new DateTime(2024, 1, 1), Assert.IsType<DateTime>(reopened.Worksheets[0]["A1"].Value));
        Assert.Equal(new DateTime(2024, 1, 2), Assert.IsType<DateTime>(reopened.Worksheets[0]["B1"].Value));
    }

    [Fact]
    public void OpenedWorkbook_WithCustomDateFormatAndOtherFormatting_PreservesThemWhenEditing()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell(
            "A1", "45292", 164, "yyyy-mm-dd", includeOtherFormatting: true);
        Assert.Empty(OoxmlValidation.ValidateXlsx(fixture));

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Worksheet sheet = workbook.Worksheets[0];
            Assert.Equal(new DateTime(2024, 1, 1), Assert.IsType<DateTime>(sheet["A1"].Value));
            Assert.Equal("yyyy-mm-dd", sheet["A1"].NumberFormat);

            sheet["A1"].SetDateTime(new DateTime(2024, 1, 2), "dd/mm/yyyy");
            produced = SaveBytes.ToBytes(workbook);
        }

        string stylesXml = TestPackage.ReadEntryText(produced, "xl/styles.xml");
        Assert.Contains("formatCode=\"yyyy-mm-dd\"", stylesXml, StringComparison.Ordinal);
        Assert.Contains("formatCode=\"dd/mm/yyyy\"", stylesXml, StringComparison.Ordinal);
        Assert.Contains("FFFFFF00", stylesXml, StringComparison.Ordinal);
        Assert.Contains("thick", stylesXml, StringComparison.Ordinal);
        Assert.Contains("center", stylesXml, StringComparison.Ordinal);

        (uint StyleIndex, uint NumberFormatId, uint FontId, uint FillId, uint BorderId, string? FormatCode) style =
            OoxmlValidation.ReadXlsxCellStyle(produced, "Hoja1", "A1");
        Assert.Equal("dd/mm/yyyy", style.FormatCode);
        Assert.Equal(1U, style.FontId);
        Assert.Equal(1U, style.FillId);
        Assert.Equal(1U, style.BorderId);

        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void OpenedWorkbook_ResettingNumberFormat_PreservesOtherFormatting()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell(
            "A1", "45292", 164, "yyyy-mm-dd", includeOtherFormatting: true);
        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            cell.NumberFormat = null;
            Assert.Equal("General", cell.NumberFormat);
            Assert.Equal(45292.0, Assert.IsType<double>(cell.Value));
            produced = SaveBytes.ToBytes(workbook);
        }

        (uint StyleIndex, uint NumberFormatId, uint FontId, uint FillId, uint BorderId, string? FormatCode) style =
            OoxmlValidation.ReadXlsxCellStyle(produced, "Hoja1", "A1");
        Assert.Equal(0U, style.NumberFormatId);
        Assert.Null(style.FormatCode);
        Assert.Equal(1U, style.FontId);
        Assert.Equal(1U, style.FillId);
        Assert.Equal(1U, style.BorderId);
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void OpenedWorkbook_WithNumberFormatInDefaultStyle_ResetsToGeneral()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell(
            "A1", "45292", 164, "yyyy-mm-dd", useDefaultStyle: true);
        Assert.Empty(OoxmlValidation.ValidateXlsx(fixture));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));
        Cell cell = workbook.Worksheets[0]["A1"];
        Assert.Equal("yyyy-mm-dd", cell.NumberFormat);
        Assert.IsType<DateTime>(cell.Value);

        cell.NumberFormat = null;

        Assert.Equal("General", cell.NumberFormat);
        Assert.Equal(45292.0, Assert.IsType<double>(cell.Value));
    }

    [Fact]
    public void Value_AssignDateTimeToLocaleSpecificBuiltInDateFormat_WritesSerial()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell("A1", "45292", 27);

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            Cell cell = workbook.Worksheets[0]["A1"];
            Assert.Null(cell.NumberFormat);
            Assert.Equal(new DateTime(2024, 1, 1), Assert.IsType<DateTime>(cell.Value));

            cell.Value = new DateTime(2024, 1, 2);
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Equal("45293", OoxmlValidation.ReadXlsxCellText(produced, "Hoja1", "A1"));
    }

    [Fact]
    public void OpenedWorkbook_WithUnknownBuiltInFormat_ReturnsTheNumberAndNullFormat()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell("A1", "1234.5", 41);

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));
        Cell cell = workbook.Worksheets[0]["A1"];

        Assert.Null(cell.NumberFormat);
        Assert.Equal(1234.5, Assert.IsType<double>(cell.Value));
    }

    [Fact]
    public void OpenedWorkbook_WithUndefinedCustomNumberFormat_ThrowsInvalidOfficeFileException()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithFormattedCell("A1", "1", 200);

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<InvalidOfficeFileException>(() => _ = sheet["A1"].Value);
        Assert.Throws<InvalidOfficeFileException>(() => _ = sheet["A1"].NumberFormat);
    }

    [Fact]
    public void DisposedWorkbook_NumberFormatAndSetDateTimeThrowObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        Cell cell = workbook.Worksheets[0]["A1"];
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = cell.NumberFormat);
        Assert.Throws<ObjectDisposedException>(() => cell.NumberFormat = "0.00");
        Assert.Throws<ObjectDisposedException>(
            () => cell.SetDateTime(new DateTime(2024, 1, 1), "yyyy-mm-dd"));
    }
}
