using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>Builds and reads the charts of a worksheet.</summary>
internal static class ExcelCharts
{
    private const uint CategoryAxisId = 111111111U;
    private const uint ValueAxisId = 222222222U;

    internal static void Build(
        ChartPart part,
        ExcelChartType type,
        string categoriesFormula,
        string valuesFormula,
        IReadOnlyList<string?> categories,
        IReadOnlyList<double?> values,
        string? seriesName,
        string? title)
    {
        C.ChartSpace space = new();
        C.Chart chart = new();
        space.AppendChild(chart);

        if (title is not null)
        {
            chart.AppendChild(new C.Title(
                new C.ChartText(new C.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(title))))),
                new C.Overlay { Val = false }));
            chart.AppendChild(new C.AutoTitleDeleted { Val = false });
        }

        C.PlotArea plotArea = new();
        chart.AppendChild(plotArea);
        plotArea.AppendChild(new C.Layout());

        C.SeriesText seriesText = new();
        if (seriesName is not null)
        {
            seriesText.AppendChild(new C.NumericValue(seriesName));
        }

        C.CategoryAxisData categoryAxisData = BuildCategories(categoriesFormula, categories);
        C.Values valuesElement = BuildValues(valuesFormula, values);

        switch (type)
        {
            case ExcelChartType.Column:
            case ExcelChartType.Bar:
            {
                C.BarChartSeries series = new(
                    new C.Index { Val = 0U },
                    new C.Order { Val = 0U });
                if (seriesName is not null)
                {
                    series.AppendChild(seriesText);
                }

                series.AppendChild(categoryAxisData);
                series.AppendChild(valuesElement);
                plotArea.AppendChild(new C.BarChart(
                    new C.BarDirection
                    {
                        Val = type == ExcelChartType.Column
                            ? C.BarDirectionValues.Column
                            : C.BarDirectionValues.Bar,
                    },
                    new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                    new C.VaryColors { Val = false },
                    series,
                    new C.AxisId { Val = CategoryAxisId },
                    new C.AxisId { Val = ValueAxisId }));
                break;
            }

            case ExcelChartType.Line:
            {
                C.LineChartSeries series = new(
                    new C.Index { Val = 0U },
                    new C.Order { Val = 0U });
                if (seriesName is not null)
                {
                    series.AppendChild(seriesText);
                }

                series.AppendChild(categoryAxisData);
                series.AppendChild(valuesElement);
                series.AppendChild(new C.Smooth { Val = false });
                plotArea.AppendChild(new C.LineChart(
                    new C.Grouping { Val = C.GroupingValues.Standard },
                    new C.VaryColors { Val = false },
                    series,
                    new C.AxisId { Val = CategoryAxisId },
                    new C.AxisId { Val = ValueAxisId }));
                break;
            }

            case ExcelChartType.Pie:
            {
                C.PieChartSeries series = new(
                    new C.Index { Val = 0U },
                    new C.Order { Val = 0U });
                if (seriesName is not null)
                {
                    series.AppendChild(seriesText);
                }

                series.AppendChild(categoryAxisData);
                series.AppendChild(valuesElement);
                plotArea.AppendChild(new C.PieChart(
                    new C.VaryColors { Val = true },
                    series));
                break;
            }

            default:
                throw new ArgumentOutOfRangeException(nameof(type), type, "The chart type is not defined.");
        }

        if (type is not ExcelChartType.Pie)
        {
            plotArea.AppendChild(new C.CategoryAxis(
                new C.AxisId { Val = CategoryAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                new C.CrossingAxis { Val = ValueAxisId }));
            plotArea.AppendChild(new C.ValueAxis(
                new C.AxisId { Val = ValueAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Left },
                new C.CrossingAxis { Val = CategoryAxisId }));
            chart.AppendChild(new C.Legend(
                new C.LegendPosition { Val = C.LegendPositionValues.Right },
                new C.Overlay { Val = false }));
        }

        chart.AppendChild(new C.PlotVisibleOnly { Val = true });
        chart.AppendChild(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });

        part.ChartSpace = space;
    }

    internal static C.CategoryAxisData BuildCategories(string formula, IReadOnlyList<string?> categories)
    {
        C.StringCache cache = new();
        cache.AppendChild(new C.PointCount { Val = (uint)categories.Count });
        for (int index = 0; index < categories.Count; index++)
        {
            if (categories[index] is not string text)
            {
                continue;
            }

            C.StringPoint point = new() { Index = (uint)index };
            point.AppendChild(new C.NumericValue(text));
            cache.AppendChild(point);
        }

        C.StringReference reference = new();
        reference.AppendChild(new C.Formula(formula));
        reference.AppendChild(cache);

        C.CategoryAxisData data = new();
        data.AppendChild(reference);
        return data;
    }

    internal static C.Values BuildValues(string formula, IReadOnlyList<double?> values)
    {
        C.NumberingCache cache = new();
        cache.AppendChild(new C.FormatCode("General"));
        cache.AppendChild(new C.PointCount { Val = (uint)values.Count });
        for (int index = 0; index < values.Count; index++)
        {
            if (values[index] is not double number)
            {
                continue;
            }

            C.NumericPoint point = new() { Index = (uint)index };
            point.AppendChild(new C.NumericValue(number.ToString("R", CultureInfo.InvariantCulture)));
            cache.AppendChild(point);
        }

        C.NumberReference reference = new();
        reference.AppendChild(new C.Formula(formula));
        reference.AppendChild(cache);

        C.Values element = new();
        element.AppendChild(reference);
        return element;
    }

    internal static Xdr.GraphicFrame BuildFrame(uint id, string name, string relationshipId, long width, long height)
    {
        Xdr.NonVisualGraphicFrameProperties nonVisual = new(
            new Xdr.NonVisualDrawingProperties { Id = id, Name = name },
            new Xdr.NonVisualGraphicFrameDrawingProperties());
        Xdr.Transform transform = new(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = width, Cy = height });
        A.GraphicData graphicData = new() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
        graphicData.AppendChild(new C.ChartReference { Id = relationshipId });
        A.Graphic graphic = new(graphicData);

        return new Xdr.GraphicFrame(nonVisual, transform, graphic) { Macro = string.Empty };
    }

    internal static List<ExcelChart> Read(DrawingsPart drawingPart)
    {
        List<ExcelChart> charts = [];
        Xdr.WorksheetDrawing? drawing = drawingPart.WorksheetDrawing;
        if (drawing is null)
        {
            return charts;
        }

        foreach (Xdr.OneCellAnchor anchor in drawing.Elements<Xdr.OneCellAnchor>())
        {
            Xdr.GraphicFrame? frame = anchor.GetFirstChild<Xdr.GraphicFrame>();
            if (frame is null)
            {
                continue;
            }

            string name = frame.NonVisualGraphicFrameProperties?
                .NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            string? relationshipId = frame.Graphic?.GraphicData?
                .GetFirstChild<C.ChartReference>()?.Id?.Value;
            if (relationshipId is null)
            {
                continue;
            }

            if (drawingPart.GetPartById(relationshipId) is not ChartPart chartPart)
            {
                continue;
            }

            Xdr.FromMarker? from = anchor.GetFirstChild<Xdr.FromMarker>();
            int column = ParseIndex(from?.ColumnId?.Text) + 1;
            int row = ParseIndex(from?.RowId?.Text) + 1;

            charts.Add(ReadChart(chartPart, name, CellReference.ToReference(column, row)));
        }

        return charts;
    }

    internal static ExcelChart ReadChart(ChartPart chartPart, string name, string reference)
    {
        C.PlotArea? plotArea = chartPart.ChartSpace?.GetFirstChild<C.Chart>()?.PlotArea;
        ExcelChartType type = ExcelChartType.Column;
        string categories = string.Empty;
        string values = string.Empty;
        string? seriesName = null;

        if (plotArea is not null)
        {
            C.BarChart? bar = plotArea.GetFirstChild<C.BarChart>();
            C.LineChart? line = plotArea.GetFirstChild<C.LineChart>();
            C.PieChart? pie = plotArea.GetFirstChild<C.PieChart>();
            OpenXmlElement? series = bar?.GetFirstChild<C.BarChartSeries>();
            series ??= line?.GetFirstChild<C.LineChartSeries>();
            series ??= pie?.GetFirstChild<C.PieChartSeries>();

            if (bar is not null)
            {
                type = bar.BarDirection?.Val?.Value == C.BarDirectionValues.Bar
                    ? ExcelChartType.Bar
                    : ExcelChartType.Column;
            }
            else if (line is not null)
            {
                type = ExcelChartType.Line;
            }
            else if (pie is not null)
            {
                type = ExcelChartType.Pie;
            }

            if (series is not null)
            {
                categories = RangeOf(series.GetFirstChild<C.CategoryAxisData>()
                    ?.GetFirstChild<C.StringReference>()?.Formula?.Text);
                values = RangeOf(series.GetFirstChild<C.Values>()
                    ?.GetFirstChild<C.NumberReference>()?.Formula?.Text);
                seriesName = series.GetFirstChild<C.SeriesText>()?.GetFirstChild<C.NumericValue>()?.Text;
            }
        }

        string? title = chartPart.ChartSpace?.GetFirstChild<C.Chart>()?.Title?
            .Descendants<A.Text>().FirstOrDefault()?.Text;

        return new ExcelChart(name, reference, type, categories, values, title, seriesName);
    }

    private static string RangeOf(string? formula)
    {
        if (string.IsNullOrEmpty(formula))
        {
            return string.Empty;
        }

        int separator = formula.LastIndexOf('!');
        return separator >= 0 ? formula[(separator + 1)..] : formula;
    }

    private static int ParseIndex(string? text)
        => int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : 0;
}
