using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>Charts anchored in a worksheet: the chart part, the drawing and their relationship.</summary>
public sealed class ExcelChartTests
{
    [Fact]
    public void AddChart_WritesTheChartPartAndTheDrawing()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A2"].Value = "Enero";
            sheet["B2"].Value = 10d;
            sheet["A3"].Value = "Febrero";
            sheet["B3"].Value = 20d;

            ExcelChart chart = sheet.AddChart(
                "D2",
                ExcelChartType.Column,
                "A2:A3",
                "B2:B3",
                title: "Ventas",
                seriesName: "Total");

            Assert.Equal("D2", chart.Reference);
            Assert.Equal(ExcelChartType.Column, chart.Type);
            Assert.Equal("A2:A3", chart.CategoriesRange);
            Assert.Equal("B2:B3", chart.ValuesRange);
            Assert.Equal("Ventas", chart.Title);
            Assert.Equal("Total", chart.SeriesName);
            Assert.Single(sheet.Charts);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Contains(TestPackage.EntryNames(bytes), name => name.EndsWith("chart1.xml", StringComparison.Ordinal));
        Assert.True(TestPackage.HasEntry(bytes, "xl/drawings/drawing1.xml"));

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        ExcelChart read = reopened.Worksheets[0].Charts[0];
        Assert.Equal("D2", read.Reference);
        Assert.Equal(ExcelChartType.Column, read.Type);
        Assert.Equal("Ventas", read.Title);
        Assert.Equal("Total", read.SeriesName);
    }

    [Theory]
    [InlineData(ExcelChartType.Column)]
    [InlineData(ExcelChartType.Bar)]
    [InlineData(ExcelChartType.Line)]
    [InlineData(ExcelChartType.Pie)]
    public void AddChart_WithEveryType_WritesAValidPackage(ExcelChartType type)
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet["A2"].Value = "Enero";
            sheet["B2"].Value = 10d;

            ExcelChart chart = sheet.AddChart("D2", type, "A2:A2", "B2:B2");

            Assert.Equal(type, chart.Type);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void AddChart_WithSeveralCharts_KeepsThemAll()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.AddChart("D2", ExcelChartType.Column, "A2:A3", "B2:B3");
        sheet.AddChart("D20", ExcelChartType.Pie, "A2:A3", "B2:B3");

        Assert.Equal(2, sheet.Charts.Count);
    }

    [Fact]
    public void AddChart_WithInvalidValues_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet.AddChart("D2:D3", ExcelChartType.Column, "A2:A3", "B2:B3"));
        Assert.Throws<ArgumentException>(() => sheet.AddChart("D2", ExcelChartType.Column, "A2", "B2:B3"));
        Assert.Throws<ArgumentException>(() => sheet.AddChart("D2", ExcelChartType.Column, "A2:A3", "B2"));
        Assert.Throws<ArgumentException>(() => sheet.AddChart("D2", ExcelChartType.Column, "A2:A3", "B2:B3", widthPoints: 0));
        Assert.Empty(sheet.Charts);
    }

    [Fact]
    public void RemoveChart_RemovesTheChart()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AddChart("D2", ExcelChartType.Column, "A2:A3", "B2:B3");

            Assert.True(sheet.RemoveChart("Gráfico 1"));
            Assert.False(sheet.RemoveChart("Gráfico 1"));
            Assert.Empty(sheet.Charts);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
    }

    [Fact]
    public void Charts_ReadsTheChartsOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithChart("Ventas", "Hoja1!$A$2:$A$4", "Hoja1!$B$2:$B$4");

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        ExcelChart chart = workbook.Worksheets[0].Charts[0];
        Assert.Equal("Ventas", chart.Name);
        Assert.Equal("A1", chart.Reference);
        Assert.Equal(ExcelChartType.Column, chart.Type);
        Assert.Equal("$A$2:$A$4", chart.CategoriesRange);
        Assert.Equal("$B$2:$B$4", chart.ValuesRange);
    }

    [Fact]
    public void Charts_ArePreservedWhenTheWorkbookIsSaved()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithChart("Ventas", "Hoja1!$A$2:$A$4", "Hoja1!$B$2:$B$4");

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Contains(TestPackage.EntryNames(produced), name => name.EndsWith("chart1.xml", StringComparison.Ordinal));
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void DisposedWorkbook_ChartsThrowObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].Charts);
    }
}
