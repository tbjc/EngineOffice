using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using CP = DocumentFormat.OpenXml.CustomProperties;
using EP = DocumentFormat.OpenXml.ExtendedProperties;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using VT = DocumentFormat.OpenXml.VariantTypes;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Tests.Support;

/// <summary>
/// Test fixtures and independent validation built directly on the Open XML SDK, so that
/// assertions never depend on the code under test.
/// </summary>
internal static class OoxmlValidation
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string EmptyCorePropertiesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
        "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" " +
        "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>";

    /// <summary>Validates a .docx package with the Open XML SDK validator.</summary>
    internal static IReadOnlyList<ValidationErrorInfo> ValidateDocx(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return new OpenXmlValidator().Validate(document).ToList();
    }

    /// <summary>Validates an .xlsx package with the Open XML SDK validator.</summary>
    internal static IReadOnlyList<ValidationErrorInfo> ValidateXlsx(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);
        return new OpenXmlValidator().Validate(document).ToList();
    }

    /// <summary>Reads the visible text of a .docx package with the Open XML SDK.</summary>
    internal static string ReadDocxText(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.Document?.Body?.InnerText ?? string.Empty;
    }

    /// <summary>Reads a cell value from a .xlsx package with the Open XML SDK.</summary>
    internal static string? ReadXlsxCellText(byte[] package, string sheetName, string cellReference)
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

        S.Cell? cell = worksheetPart.Worksheet!.Descendants<S.Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.Ordinal));
        if (cell is null)
        {
            return null;
        }

        if (cell.DataType?.Value == S.CellValues.InlineString)
        {
            return cell.InlineString?.InnerText;
        }

        if (cell.DataType?.Value == S.CellValues.SharedString)
        {
            int index = int.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture);
            return workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<S.SharedStringItem>().ElementAt(index).InnerText;
        }

        return cell.CellValue?.Text;
    }

    /// <summary>Builds an .xlsx whose worksheet declares data validation rules.</summary>
    internal static byte[] CreateXlsxWithDataValidations(
        params (string Type, string Sqref, string? Formula1, string? Formula2)[] validations)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        S.Worksheet worksheet = new(new S.SheetData());
        if (validations.Length > 0)
        {
            S.DataValidations container = new();
            foreach ((string type, string sqref, string? formula1, string? formula2) in validations)
            {
                S.DataValidation element = new()
                {
                    Type = new EnumValue<S.DataValidationValues>(
                        new S.DataValidationValues(type)),
                    SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue(sqref) }),
                };
                if (formula1 is not null)
                {
                    element.Formula1 = new S.Formula1(formula1);
                }

                if (formula2 is not null)
                {
                    element.Formula2 = new S.Formula2(formula2);
                }

                container.AppendChild(element);
            }

            worksheet.AppendChild(container);
        }

        worksheetPart.Worksheet = worksheet;

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Reads the data validation rules of a worksheet with the Open XML SDK.</summary>
    internal static IReadOnlyList<(string? Type, string? Sqref, string? Formula1, string? Formula2, string? Operator)>
        ReadXlsxDataValidations(byte[] package, string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

        S.DataValidations? container = worksheetPart.Worksheet!.GetFirstChild<S.DataValidations>();
        if (container is null)
        {
            return [];
        }

        return container.Elements<S.DataValidation>()
            .Select(element => (
                element.Type?.InnerText,
                element.SequenceOfReferences?.InnerText,
                element.Formula1?.Text,
                element.Formula2?.Text,
                element.Operator?.InnerText))
            .ToList();
    }

    /// <summary>Builds an .xlsx whose worksheet declares an auto filter.</summary>
    internal static byte[] CreateXlsxWithAutoFilter(string reference)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(
            new S.SheetData(),
            new S.AutoFilter { Reference = reference });

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Reads the auto filter range of a worksheet with the Open XML SDK.</summary>
    internal static string? ReadXlsxAutoFilter(byte[] package, string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

        return worksheetPart.Worksheet!.GetFirstChild<S.AutoFilter>()?.Reference?.Value;
    }

    /// <summary>Builds an .xlsx whose worksheet declares a structured table.</summary>
    internal static byte[] CreateXlsxWithTable(
        string name,
        string reference,
        string[] columnNames,
        bool hasHeaderRow = true)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());

        TableDefinitionPart tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
        S.Table table = new()
        {
            Id = 1U,
            Name = name,
            DisplayName = name,
            Reference = reference,
            HeaderRowCount = hasHeaderRow ? 1U : 0U,
        };
        table.AppendChild(new S.AutoFilter { Reference = reference });
        S.TableColumns columns = new() { Count = (uint)columnNames.Length };
        for (int index = 0; index < columnNames.Length; index++)
        {
            columns.AppendChild(new S.TableColumn { Id = (uint)(index + 1), Name = columnNames[index] });
        }

        table.AppendChild(columns);
        tablePart.Table = table;

        worksheetPart.Worksheet.AppendChild(new S.TableParts(
            new S.TablePart { Id = worksheetPart.GetIdOfPart(tablePart) })
        { Count = 1U });

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Reads the structured tables of a worksheet with the Open XML SDK.</summary>
    internal static IReadOnlyList<(string Name, string? Reference, uint HeaderRowCount, string[] Columns)>
        ReadXlsxTables(byte[] package, string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

        List<(string Name, string? Reference, uint HeaderRowCount, string[] Columns)> tables = [];
        foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts)
        {
            S.Table? table = tablePart.Table;
            if (table is null)
            {
                continue;
            }

            string[] columns = table.TableColumns is null
                ? []
                : table.TableColumns.Elements<S.TableColumn>()
                    .Select(column => column.Name?.Value ?? string.Empty)
                    .ToArray();
            tables.Add((
                table.DisplayName?.Value ?? table.Name?.Value ?? string.Empty,
                table.Reference?.Value,
                table.HeaderRowCount?.Value ?? 1U,
                columns));
        }

        return tables;
    }

    /// <summary>Builds an .xlsx whose worksheet holds a column chart.</summary>
    internal static byte[] CreateXlsxWithChart(string name, string categoriesFormula, string valuesFormula)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());

        DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();

        C.ChartSpace space = new();
        C.Chart chart = new();
        space.AppendChild(chart);
        C.PlotArea plotArea = new(new C.Layout());
        chart.AppendChild(plotArea);
        plotArea.AppendChild(new C.BarChart(
            new C.BarDirection { Val = C.BarDirectionValues.Column },
            new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
            new C.BarChartSeries(
                new C.Index { Val = 0U },
                new C.Order { Val = 0U },
                new C.CategoryAxisData(new C.StringReference(new C.Formula(categoriesFormula))),
                new C.Values(new C.NumberReference(new C.Formula(valuesFormula)))),
            new C.AxisId { Val = 111111111U },
            new C.AxisId { Val = 222222222U }));
        plotArea.AppendChild(new C.CategoryAxis(
            new C.AxisId { Val = 111111111U },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.CrossingAxis { Val = 222222222U }));
        plotArea.AppendChild(new C.ValueAxis(
            new C.AxisId { Val = 222222222U },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Left },
            new C.CrossingAxis { Val = 111111111U }));
        chart.AppendChild(new C.PlotVisibleOnly { Val = true });
        chartPart.ChartSpace = space;

        Xdr.WorksheetDrawing drawing = new();
        Xdr.GraphicFrame frame = new(
            new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = 2U, Name = name },
                new Xdr.NonVisualGraphicFrameDrawingProperties()),
            new Xdr.Transform(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 4572000L, Cy = 2743200L }),
            new A.Graphic(new A.GraphicData(new C.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) })
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart",
            }))
        { Macro = string.Empty };
        drawing.AppendChild(new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId("0"),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId("0"),
                new Xdr.RowOffset("0")),
            new Xdr.Extent { Cx = 4572000L, Cy = 2743200L },
            frame,
            new Xdr.ClientData()));
        drawingsPart.WorksheetDrawing = drawing;

        worksheetPart.Worksheet.AppendChild(new S.Drawing
        {
            Id = worksheetPart.GetIdOfPart(drawingsPart),
        });

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Builds an .xlsx whose worksheet declares an image anchored in a cell.</summary>
    internal static byte[] CreateXlsxWithImage(string name, string reference, byte[] content)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());

        DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
        using (Stream imageStream = imagePart.GetStream(FileMode.Create, FileAccess.Write))
        {
            imageStream.Write(content, 0, content.Length);
        }

        Xdr.WorksheetDrawing drawing = new();
        Xdr.FromMarker from = new(
            new Xdr.ColumnId("0"),
            new Xdr.ColumnOffset("0"),
            new Xdr.RowId("0"),
            new Xdr.RowOffset("0"));
        Xdr.Extent extent = new() { Cx = 1270000L, Cy = 635000L };
        A.Blip blip = new() { Embed = drawingsPart.GetIdOfPart(imagePart) };
        Xdr.BlipFill blipFill = new(blip, new A.Stretch(new A.FillRectangle()));
        A.Transform2D transform = new(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = 1270000L, Cy = 635000L });
        A.PresetGeometry geometry = new(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle };
        Xdr.Picture picture = new(
            new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties { Id = 2U, Name = name },
                new Xdr.NonVisualPictureDrawingProperties()),
            blipFill,
            new Xdr.ShapeProperties(transform, geometry));
        drawing.AppendChild(new Xdr.OneCellAnchor(from, extent, picture, new Xdr.ClientData()));
        drawingsPart.WorksheetDrawing = drawing;

        worksheetPart.Worksheet.AppendChild(new S.Drawing
        {
            Id = worksheetPart.GetIdOfPart(drawingsPart),
        });

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Builds an .xlsx whose worksheet declares notes.</summary>
    internal static byte[] CreateXlsxWithComments(params (string Reference, string Text, string Author)[] comments)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());

        if (comments.Length > 0)
        {
            List<string> authors = [];
            foreach ((string _, string _, string author) in comments)
            {
                if (!authors.Contains(author, StringComparer.Ordinal))
                {
                    authors.Add(author);
                }
            }

            S.Authors authorsElement = new();
            foreach (string author in authors)
            {
                authorsElement.AppendChild(new S.Author(author));
            }

            S.CommentList commentList = new();
            foreach ((string reference, string text, string author) in comments)
            {
                commentList.AppendChild(new S.Comment
                {
                    Reference = reference,
                    AuthorId = (uint)authors.IndexOf(author),
                    CommentText = new S.CommentText(new S.Text(text)),
                });
            }

            WorksheetCommentsPart commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
            commentsPart.Comments = new S.Comments(authorsElement, commentList);

            VmlDrawingPart drawing = worksheetPart.AddNewPart<VmlDrawingPart>();
            using (Stream drawingStream = drawing.GetStream(FileMode.Create, FileAccess.Write))
            using (StreamWriter writer = new(drawingStream, new UTF8Encoding(false)))
            {
                writer.Write(
                    "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                    "xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
                    "xmlns:x=\"urn:schemas-microsoft-com:office:excel\">" +
                    "<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>" +
                    "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" " +
                    "path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/>" +
                    "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>" +
                    "<v:shape id=\"_x0000_s1025\" type=\"#_x0000_t202\" " +
                    "style=\"position:absolute;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">" +
                    "<v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox>" +
                    "<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/>" +
                    "<x:AutoFill>False</x:AutoFill><x:Row>0</x:Row><x:Column>0</x:Column>" +
                    "</x:ClientData></v:shape></xml>");
            }

            worksheetPart.Worksheet.AppendChild(new S.LegacyDrawing
            {
                Id = worksheetPart.GetIdOfPart(drawing),
            });
        }

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Reads the notes of a worksheet with the Open XML SDK.</summary>
    internal static IReadOnlyList<(string Reference, string Text, string Author)> ReadXlsxComments(
        byte[] package,
        string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorksheetPart worksheetPart = WorksheetPart(document, sheetName);
        S.Comments? comments = worksheetPart.WorksheetCommentsPart?.Comments;
        if (comments is null)
        {
            return [];
        }

        List<string> authors = [];
        if (comments.Authors is not null)
        {
            foreach (S.Author author in comments.Authors.Elements<S.Author>())
            {
                authors.Add(author.Text);
            }
        }

        List<(string Reference, string Text, string Author)> result = [];
        if (comments.CommentList is null)
        {
            return result;
        }

        foreach (S.Comment comment in comments.CommentList.Elements<S.Comment>())
        {
            uint authorId = comment.AuthorId?.Value ?? 0U;
            result.Add((
                comment.Reference?.Value ?? string.Empty,
                comment.CommentText?.InnerText ?? string.Empty,
                authorId < authors.Count ? authors[(int)authorId] : string.Empty));
        }

        return result;
    }

    /// <summary>Builds an .xlsx whose worksheet declares hyperlinks.</summary>
    internal static byte[] CreateXlsxWithHyperlinks(params (string Reference, string? Target, string? Location)[] links)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        S.Worksheet worksheet = new(new S.SheetData());

        if (links.Length > 0)
        {
            S.Hyperlinks hyperlinks = new();
            foreach ((string reference, string? target, string? location) in links)
            {
                S.Hyperlink element = new() { Reference = reference };
                if (target is not null)
                {
                    HyperlinkRelationship relationship = worksheetPart.AddHyperlinkRelationship(
                        new Uri(target, UriKind.Absolute),
                        true);
                    element.Id = relationship.Id;
                }

                if (location is not null)
                {
                    element.Location = location;
                }

                hyperlinks.AppendChild(element);
            }

            worksheet.AppendChild(hyperlinks);
        }

        worksheetPart.Worksheet = worksheet;

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Reads the hyperlinks of a worksheet with the Open XML SDK.</summary>
    internal static IReadOnlyList<(string Reference, string? Target, string? Location)> ReadXlsxHyperlinks(
        byte[] package,
        string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorksheetPart worksheetPart = WorksheetPart(document, sheetName);
        S.Hyperlinks? hyperlinks = worksheetPart.Worksheet!.GetFirstChild<S.Hyperlinks>();
        if (hyperlinks is null)
        {
            return [];
        }

        List<(string Reference, string? Target, string? Location)> result = [];
        foreach (S.Hyperlink element in hyperlinks.Elements<S.Hyperlink>())
        {
            string? target = null;
            if (element.Id?.Value is string relationshipId)
            {
                target = worksheetPart.HyperlinkRelationships
                    .FirstOrDefault(relationship => relationship.Id == relationshipId)?
                    .Uri
                    .ToString();
            }

            result.Add((element.Reference?.Value ?? string.Empty, target, element.Location?.Value));
        }

        return result;
    }

    /// <summary>Builds an .xlsx whose worksheet declares print options, margins and page setup.</summary>
    internal static byte[] CreateXlsxWithPrintSettings(
        double left = 0.7,
        double right = 0.7,
        double top = 0.75,
        double bottom = 0.75,
        double header = 0.3,
        double footer = 0.3,
        string orientation = "portrait",
        uint paperSize = 9U,
        uint? scale = null,
        uint? fitToWidth = null,
        uint? fitToHeight = null,
        bool horizontalCentered = false,
        bool verticalCentered = false,
        bool gridLines = false,
        bool headings = false)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        S.Worksheet worksheet = new();
        worksheet.AppendChild(new S.PrintOptions
        {
            HorizontalCentered = horizontalCentered,
            VerticalCentered = verticalCentered,
            GridLines = gridLines,
            Headings = headings,
        });
        worksheet.AppendChild(new S.PageMargins
        {
            Left = left,
            Right = right,
            Top = top,
            Bottom = bottom,
            Header = header,
            Footer = footer,
        });
        worksheet.AppendChild(new S.PageSetup
        {
            PaperSize = paperSize,
            Orientation = new S.OrientationValues(orientation),
            Scale = scale,
            FitToWidth = fitToWidth,
            FitToHeight = fitToHeight,
        });
        worksheet.AppendChild(new S.SheetData());
        worksheetPart.Worksheet = worksheet;

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Reads the page margins of a worksheet with the Open XML SDK.</summary>
    internal static (double Left, double Right, double Top, double Bottom, double Header, double Footer)?
        ReadXlsxPageMargins(byte[] package, string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        S.PageMargins? margins = WorksheetPart(document, sheetName).Worksheet!.GetFirstChild<S.PageMargins>();
        if (margins is null)
        {
            return null;
        }

        return (
            margins.Left?.Value ?? 0d,
            margins.Right?.Value ?? 0d,
            margins.Top?.Value ?? 0d,
            margins.Bottom?.Value ?? 0d,
            margins.Header?.Value ?? 0d,
            margins.Footer?.Value ?? 0d);
    }

    /// <summary>Reads the page setup of a worksheet with the Open XML SDK.</summary>
    internal static (uint PaperSize, string Orientation, uint? Scale, uint? FitToWidth, uint? FitToHeight)?
        ReadXlsxPageSetup(byte[] package, string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        S.PageSetup? setup = WorksheetPart(document, sheetName).Worksheet!.GetFirstChild<S.PageSetup>();
        if (setup is null)
        {
            return null;
        }

        return (
            setup.PaperSize?.Value ?? 0U,
            setup.Orientation?.InnerText ?? string.Empty,
            setup.Scale?.Value,
            setup.FitToWidth?.Value,
            setup.FitToHeight?.Value);
    }

    /// <summary>Reads the print options of a worksheet with the Open XML SDK.</summary>
    internal static (bool HorizontalCentered, bool VerticalCentered, bool GridLines, bool Headings)?
        ReadXlsxPrintOptions(byte[] package, string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        S.PrintOptions? options = WorksheetPart(document, sheetName).Worksheet!.GetFirstChild<S.PrintOptions>();
        if (options is null)
        {
            return null;
        }

        return (
            options.HorizontalCentered?.Value ?? false,
            options.VerticalCentered?.Value ?? false,
            options.GridLines?.Value ?? false,
            options.Headings?.Value ?? false);
    }

    private static WorksheetPart WorksheetPart(SpreadsheetDocument document, string sheetName)
    {
        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }

    /// <summary>Reads the shared string table of a workbook with the Open XML SDK.</summary>
    internal static (IReadOnlyList<string> Items, uint? Count, uint? UniqueCount) ReadXlsxSharedStrings(
        byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        S.SharedStringTable? table = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
        if (table is null)
        {
            return ([], null, null);
        }

        List<string> items = [];
        foreach (S.SharedStringItem item in table.Elements<S.SharedStringItem>())
        {
            items.Add(item.InnerText);
        }

        return (items, table.Count?.Value, table.UniqueCount?.Value);
    }

    /// <summary>Builds an .xlsx whose worksheet declares a frozen pane, a row height and a column width.</summary>
    internal static byte[] CreateXlsxWithSheetLayout(
        int frozenColumns = 0,
        int frozenRows = 0,
        string topLeftCell = "A1",
        int rowIndex = 1,
        double rowHeight = 0,
        bool rowHidden = false,
        int columnIndex = 1,
        double columnWidth = 0,
        bool columnHidden = false,
        int columnMax = 0)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        S.Worksheet worksheet = new();

        S.SheetView view = new() { WorkbookViewId = 0U };
        if (frozenRows > 0 || frozenColumns > 0)
        {
            view.AppendChild(new S.Pane
            {
                HorizontalSplit = frozenColumns > 0 ? frozenColumns : null,
                VerticalSplit = frozenRows > 0 ? frozenRows : null,
                TopLeftCell = topLeftCell,
                ActivePane = new S.PaneValues("bottomRight"),
                State = S.PaneStateValues.Frozen,
            });
        }

        worksheet.AppendChild(new S.SheetViews(view));

        if (columnWidth > 0 || columnHidden)
        {
            worksheet.AppendChild(new S.Columns(new S.Column
            {
                Min = (uint)columnIndex,
                Max = (uint)(columnMax > 0 ? columnMax : columnIndex),
                Width = columnWidth > 0 ? columnWidth : null,
                CustomWidth = columnWidth > 0,
                Hidden = columnHidden,
            }));
        }

        S.Row row = new(new S.Cell { CellReference = "A1", CellValue = new S.CellValue("x") })
        {
            RowIndex = (uint)rowIndex,
            Height = rowHeight > 0 ? rowHeight : null,
            CustomHeight = rowHeight > 0,
            Hidden = rowHidden,
        };
        worksheet.AppendChild(new S.SheetData(row));
        worksheetPart.Worksheet = worksheet;

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Reads the frozen pane of a worksheet with the Open XML SDK.</summary>
    internal static (double? XSplit, double? YSplit, string? TopLeftCell, string? State, string? ActivePane)? ReadXlsxPane(
        byte[] package,
        string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

        S.Pane? pane = worksheetPart.Worksheet!.Descendants<S.Pane>().FirstOrDefault();
        if (pane is null)
        {
            return null;
        }

        return (pane.HorizontalSplit?.Value, pane.VerticalSplit?.Value, pane.TopLeftCell?.Value, pane.State?.InnerText, pane.ActivePane?.InnerText);
    }

    /// <summary>Reads the attributes of a row with the Open XML SDK.</summary>
    internal static (double? Height, bool Hidden) ReadXlsxRow(
        byte[] package,
        uint rowIndex,
        string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

        S.Row? row = worksheetPart.Worksheet!.Descendants<S.Row>()
            .FirstOrDefault(candidate => candidate.RowIndex?.Value == rowIndex);
        return (row?.Height?.Value, row?.Hidden?.Value ?? false);
    }

    /// <summary>Reads the attributes of a column with the Open XML SDK.</summary>
    internal static (double? Width, bool Hidden) ReadXlsxColumn(
        byte[] package,
        int columnIndex,
        string sheetName = "Hoja1")
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

        S.Column? column = worksheetPart.Worksheet!.Descendants<S.Column>()
            .FirstOrDefault(candidate => candidate.Min?.Value <= columnIndex && candidate.Max?.Value >= columnIndex);
        return (column?.Width?.Value, column?.Hidden?.Value ?? false);
    }

    /// <summary>Builds an .xlsx whose workbook declares defined names.</summary>
    internal static byte[] CreateXlsxWithDefinedNames(params (string Name, string Text, uint? LocalSheetId)[] names)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        if (names.Length > 0)
        {
            S.DefinedNames definedNames = new();
            foreach ((string name, string text, uint? localSheetId) in names)
            {
                S.DefinedName element = new() { Name = name };
                if (localSheetId is not null)
                {
                    element.LocalSheetId = localSheetId;
                }

                element.Text = text;
                definedNames.AppendChild(element);
            }

            workbookPart.Workbook.AppendChild(definedNames);
        }

        return Flush(document, stream);
    }

    /// <summary>Reads the defined names of a workbook with the Open XML SDK.</summary>
    internal static IReadOnlyList<(string Name, string Text, uint? LocalSheetId)> ReadXlsxDefinedNames(
        byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        S.DefinedNames? names = document.WorkbookPart?.Workbook?.GetFirstChild<S.DefinedNames>();
        if (names is null)
        {
            return [];
        }

        return names.Elements<S.DefinedName>()
            .Select(element => (element.Name?.Value ?? string.Empty, element.Text, element.LocalSheetId?.Value))
            .ToList();
    }

    /// <summary>Reads a cell formula and its cached value from a .xlsx package with the SDK.</summary>
    internal static (string? Formula, string? CachedValue) ReadXlsxCellFormula(
        byte[] package,
        string sheetName,
        string cellReference)
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

        S.Cell? cell = worksheetPart.Worksheet!.Descendants<S.Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.Ordinal));
        return (cell?.CellFormula?.Text, cell?.CellValue?.Text);
    }

    /// <summary>Reads the workbook's full-calculation-on-load flag with the SDK.</summary>
    internal static bool? ReadXlsxFullCalculationOnLoad(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);
        return document.WorkbookPart?.Workbook?.CalculationProperties?.FullCalculationOnLoad?.Value;
    }

    /// <summary>
    /// Reads the properties of the first run in a .docx package with the Open XML SDK, so tests
    /// can assert on formatting without going through the code under test.
    /// </summary>
    internal static W.RunProperties? ReadDocxFirstRunProperties(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.Document?.Body?.Descendants<W.Run>().FirstOrDefault()?.RunProperties;
    }

    /// <summary>
    /// Builds a .docx whose single run carries full character formatting, including a property
    /// OfficeEngine does not model (<c>w:highlight</c>), in the child order required by the schema.
    /// </summary>
    internal static byte[] CreateDocxWithFormattedRun(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(
                new W.RunProperties(
                    new W.RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" },
                    new W.Bold(),
                    new W.Italic(),
                    new W.Color { Val = "00FF00" },
                    new W.FontSize { Val = "28" },
                    new W.Highlight { Val = W.HighlightColorValues.Yellow },
                    new W.Underline { Val = W.UnderlineValues.Double }),
                new W.Text(text)))));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose single paragraph carries the given paragraph properties. It lets tests
    /// place exactly the markup they need (including values OfficeEngine does not model) without
    /// going through the code under test.
    /// </summary>
    internal static byte[] CreateDocxWithParagraphProperties(string text, W.ParagraphProperties properties)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(new W.Paragraph(properties, new W.Run(new W.Text(text)))));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose single paragraph stores the given raw justification value, so tests can
    /// cover a value OfficeEngine models, one it does not model, and one that cannot be parsed.
    /// </summary>
    internal static byte[] CreateDocxWithStoredAlignment(string text, string value)
    {
        W.Justification justification = new();
        justification.SetAttribute(new OpenXmlAttribute("w", "val", WordprocessingNamespace, value));
        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(justification));
    }

    /// <summary>
    /// Builds a .docx whose single paragraph stores raw line-spacing attributes, for example the
    /// multiple-of-a-line form (<c>w:lineRule="auto"</c>) that cannot be expressed in points.
    /// </summary>
    internal static byte[] CreateDocxWithStoredLineSpacing(string text, string line, string? lineRule)
    {
        W.SpacingBetweenLines spacing = new();
        spacing.SetAttribute(new OpenXmlAttribute("w", "line", WordprocessingNamespace, line));
        if (lineRule is not null)
        {
            spacing.SetAttribute(new OpenXmlAttribute("w", "lineRule", WordprocessingNamespace, lineRule));
        }

        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(spacing));
    }

    /// <summary>
    /// Builds a .docx whose single paragraph stores the given raw spacing-before value, so tests can
    /// cover a malformed value that cannot be parsed.
    /// </summary>
    internal static byte[] CreateDocxWithStoredSpacing(string text, string before)
    {
        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(
            new W.SpacingBetweenLines { Before = before }));
    }

    /// <summary>
    /// Builds a .docx whose single paragraph stores the given raw indent attributes, including the
    /// character-based variant (<c>w:leftChars</c>) that OfficeEngine preserves but does not model.
    /// </summary>
    internal static byte[] CreateDocxWithStoredIndent(string text, int? leftChars, string? firstLine)
    {
        W.Indentation indentation = new();
        if (leftChars is not null)
        {
            indentation.LeftChars = leftChars.Value;
        }

        if (firstLine is not null)
        {
            indentation.FirstLine = firstLine;
        }

        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(indentation));
    }

    /// <summary>
    /// Builds a .docx whose single paragraph mixes formatting OfficeEngine models (alignment,
    /// indentation, spacing) with markup it does not (keep-with-next), in schema order.
    /// </summary>
    internal static byte[] CreateDocxWithMixedParagraphFormatting(string text)
    {
        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(
            new W.KeepNext(),
            new W.SpacingBetweenLines { Before = "120" },
            new W.Indentation { Left = "360", FirstLine = "240" },
            new W.Justification { Val = W.JustificationValues.Center }));
    }

    /// <summary>
    /// Builds a .docx whose single paragraph carries numbering properties (<c>w:numPr</c>), a child
    /// of <c>w:pPr</c> this version does not model.
    /// </summary>
    internal static byte[] CreateDocxWithNumbering(string text)
    {
        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(
            new W.NumberingProperties(
                new W.NumberingLevelReference { Val = 0 },
                new W.NumberingId { Val = 1 })));
    }

    /// <summary>
    /// Builds a .docx whose single paragraph carries layout flags this version does not model
    /// (<c>w:overflowPunct</c>, <c>w:bidi</c>, <c>w:adjustRightInd</c>, <c>w:snapToGrid</c>), so
    /// tests can check that the properties added by the library are still inserted in schema order.
    /// </summary>
    internal static byte[] CreateDocxWithUnmodelledParagraphFlags(string text)
    {
        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(
            new W.OverflowPunctuation { Val = false },
            new W.BiDi(),
            new W.AdjustRightIndent { Val = false },
            new W.SnapToGrid { Val = false }));
    }

    /// <summary>
    /// Reads the paragraph properties of the first paragraph in a .docx package with the Open XML
    /// SDK, so tests can assert on paragraph formatting without going through the code under test.
    /// </summary>
    internal static W.ParagraphProperties? ReadDocxFirstParagraphProperties(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.Document?.Body?.Descendants<W.Paragraph>().FirstOrDefault()?.ParagraphProperties;
    }

    /// <summary>
    /// Builds a .docx whose run colour comes from the theme rather than a fixed RGB value.
    /// </summary>
    internal static byte[] CreateDocxWithThemeColor(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(
                new W.RunProperties(
                    new W.Color { Val = "auto", ThemeColor = W.ThemeColorValues.Accent1, ThemeShade = "80" }),
                new W.Text(text)))));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx with a paragraph property (justification) that Phase 1 and Phase 2 do not model.
    /// </summary>
    internal static byte[] CreateDocxWithCenteredParagraph(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(
                new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center }),
                new W.Run(new W.Text(text)))));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose run carries a font size that is not a number, to exercise the
    /// error path OfficeEngine must report instead of guessing.
    /// </summary>
    internal static byte[] CreateDocxWithMalformedFontSize(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(
                new W.RunProperties(new W.FontSize { Val = "not-a-number" }),
                new W.Text(text)))));

        return Flush(document, stream);
    }

    /// <summary>Builds a .docx that contains a paragraph, a table and another paragraph.</summary>
    internal static byte[] CreateDocxWithTable(string before, string cellText, string after)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(new W.Text(before))),
            new W.Table(
                new W.TableProperties(new W.TableWidth { Width = "0", Type = W.TableWidthUnitValues.Auto }),
                new W.TableGrid(new W.GridColumn { Width = "2400" }),
                new W.TableRow(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(cellText)))))),
            new W.Paragraph(new W.Run(new W.Text(after)))));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose body holds the given children in order, so tests can place exactly the
    /// block-level markup they need (paragraphs, tables and content this version does not model)
    /// without going through the code under test.
    /// </summary>
    internal static byte[] CreateDocxWithBody(params OpenXmlElement[] bodyChildren)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(bodyChildren));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose single section references the requested headers and footers. A
    /// <see langword="null"/> text omits that reference, so the helper can produce any combination of
    /// types. The part named by each reference holds the text in a single paragraph.
    /// </summary>
    internal static byte[] CreateDocxWithHeaders(
        string bodyText,
        string? primaryHeader = null,
        string? firstHeader = null,
        string? evenHeader = null,
        string? primaryFooter = null,
        string? firstFooter = null,
        string? evenFooter = null,
        bool differentFirstPage = false)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();

        List<OpenXmlElement> properties = [];
        void AddReference(bool header, W.HeaderFooterValues type, string? text)
        {
            if (text is null)
            {
                return;
            }

            if (header)
            {
                HeaderPart part = mainPart.AddNewPart<HeaderPart>();
                part.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text(text))));
                properties.Add(new W.HeaderReference { Type = type, Id = mainPart.GetIdOfPart(part) });
            }
            else
            {
                FooterPart part = mainPart.AddNewPart<FooterPart>();
                part.Footer = new W.Footer(new W.Paragraph(new W.Run(new W.Text(text))));
                properties.Add(new W.FooterReference { Type = type, Id = mainPart.GetIdOfPart(part) });
            }
        }

        AddReference(header: true, W.HeaderFooterValues.Default, primaryHeader);
        AddReference(header: true, W.HeaderFooterValues.First, firstHeader);
        AddReference(header: true, W.HeaderFooterValues.Even, evenHeader);
        AddReference(header: false, W.HeaderFooterValues.Default, primaryFooter);
        AddReference(header: false, W.HeaderFooterValues.First, firstFooter);
        AddReference(header: false, W.HeaderFooterValues.Even, evenFooter);
        if (differentFirstPage)
        {
            properties.Add(new W.TitlePage());
        }

        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(new W.Text(bodyText))),
            new W.SectionProperties(properties)));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx with two sections that reference the same header part, so tests can check that
    /// editing the shared part through one section is visible through the other.
    /// </summary>
    internal static byte[] CreateDocxWithTwoSectionsSharingAHeader(
        string firstBodyText,
        string secondBodyText,
        string headerText)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();

        HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text(headerText))));
        string id = mainPart.GetIdOfPart(headerPart);

        W.SectionProperties firstSection = new(new W.HeaderReference
        {
            Type = W.HeaderFooterValues.Default,
            Id = id,
        });
        W.SectionProperties secondSection = new(new W.HeaderReference
        {
            Type = W.HeaderFooterValues.Default,
            Id = id,
        });

        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.ParagraphProperties(firstSection), new W.Run(new W.Text(firstBodyText))),
            new W.Paragraph(new W.Run(new W.Text(secondBodyText))),
            secondSection));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose body does not reference the header part it contains, so tests can check
    /// that an unreferenced part is preserved instead of being dropped.
    /// </summary>
    internal static byte[] CreateDocxWithOrphanHeader(string bodyText, string headerText)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text(headerText))));
        mainPart.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text(bodyText)))));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose section references a footer part as if it were a header, so tests can check
    /// that a reference whose relationship is the wrong kind fails instead of being read as an empty
    /// header.
    /// </summary>
    internal static byte[] CreateDocxWithHeaderReferenceToFooterPart(string bodyText)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
        footerPart.Footer = new W.Footer(new W.Paragraph(new W.Run(new W.Text("pie"))));
        W.SectionProperties section = new(new W.HeaderReference
        {
            Type = W.HeaderFooterValues.Default,
            Id = mainPart.GetIdOfPart(footerPart),
        });
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(new W.Text(bodyText))),
            section));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose header holds a paragraph and a content control (<c>w:sdt</c>, which this
    /// version does not model), so tests can check that editing the header text preserves it.
    /// </summary>
    internal static byte[] CreateDocxWithUnmodelledHeaderContent(string bodyText, string headerText)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new W.Header(
            new W.Paragraph(new W.Run(new W.Text(headerText))),
            new W.SdtBlock(new W.SdtContentBlock(new W.Paragraph(new W.Run(new W.Text("dentro"))))));
        W.SectionProperties section = new(new W.HeaderReference
        {
            Type = W.HeaderFooterValues.Default,
            Id = mainPart.GetIdOfPart(headerPart),
        });
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(new W.Text(bodyText))),
            section));

        return Flush(document, stream);
    }

    /// <summary>A 1×1 transparent PNG used by image fixtures and tests.</summary>
    internal static byte[] TinyPngBytes { get; } = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    /// <summary>
    /// Builds a .docx with a text paragraph and a paragraph that embeds a PNG through an inline
    /// drawing, using the SDK instead of the code under test.
    /// </summary>
    internal static byte[] CreateDocxWithInlineImage(
        string bodyText,
        long widthEmus = 12700,
        long heightEmus = 12700)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using MemoryStream data = new(TinyPngBytes, writable: false);
        imagePart.FeedData(data);
        string id = mainPart.GetIdOfPart(imagePart);

        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(new W.Text(bodyText))),
            CreateParagraphWithInlineImage(id, 1U, widthEmus, heightEmus)));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx with an inline picture that has no <c>wp:extent</c>, so tests can check that the
    /// reader rejects an image whose size is unknown instead of inventing one.
    /// </summary>
    internal static byte[] CreateDocxWithInlineImageWithoutExtent(string bodyText)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using MemoryStream data = new(TinyPngBytes, writable: false);
        imagePart.FeedData(data);
        string id = mainPart.GetIdOfPart(imagePart);

        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(new W.Text(bodyText))),
            CreateParagraphWithInlineImage(id, 1U, 12700, 12700, includeExtent: false)));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose inline picture points at a part that is not an image, so tests can check
    /// that the reader rejects the relationship instead of treating any part as an image.
    /// </summary>
    internal static byte[] CreateDocxWithInlineImagePointingToANonImagePart(string bodyText)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles();
        string id = mainPart.GetIdOfPart(stylesPart);

        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(new W.Text(bodyText))),
            CreateParagraphWithInlineImage(id, 1U, 12700, 12700)));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose inline picture names an image relationship that the package does not
    /// contain, so tests can check the typed failure instead of an empty image list.
    /// </summary>
    internal static byte[] CreateDocxWithInlineImageReferenceToMissingPart(string bodyText)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(new W.Text(bodyText))),
            CreateParagraphWithInlineImage("rId999", 1U, 12700, 12700)));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx with a hyperlink whose destination is an external relationship, so tests can
    /// check that a relationship this version does not model survives an edit.
    /// </summary>
    internal static byte[] CreateDocxWithHyperlink(string text, string uri)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        string id = mainPart.AddHyperlinkRelationship(new Uri(uri), isExternal: true).Id;
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Hyperlink(new W.Run(new W.Text(text))) { Id = id })));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx with two hyperlinks that share one external relationship, so tests can check
    /// that editing one link does not change the other.
    /// </summary>
    internal static byte[] CreateDocxWithTwoHyperlinksSharingRelationship(string firstText, string secondText, string uri)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        string id = mainPart.AddHyperlinkRelationship(new Uri(uri), isExternal: true).Id;
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(
                new W.Hyperlink(new W.Run(new W.Text(firstText))) { Id = id },
                new W.Hyperlink(new W.Run(new W.Text(secondText))) { Id = id })));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx with an internal hyperlink to a bookmark, so tests can check that a link this
    /// version does not model is preserved instead of being reported as an external one.
    /// </summary>
    internal static byte[] CreateDocxWithInternalHyperlink(string text, string anchor)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Hyperlink(new W.Run(new W.Text(text))) { Anchor = anchor })));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose hyperlink names a relationship the package does not contain, so tests
    /// can check the typed failure instead of an empty list.
    /// </summary>
    internal static byte[] CreateDocxWithHyperlinkReferenceToMissingRelationship(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Hyperlink(new W.Run(new W.Text(text))) { Id = "rId999" })));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose hyperlink names an internal hyperlink relationship, so tests can check
    /// that a relationship which is not external is rejected instead of being reported as a URL.
    /// </summary>
    internal static byte[] CreateDocxWithNonExternalHyperlinkRelationship(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        HyperlinkRelationship relationship = mainPart.AddHyperlinkRelationship(
            new Uri("foo/bar", UriKind.Relative),
            isExternal: false);
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Hyperlink(new W.Run(new W.Text(text))) { Id = relationship.Id })));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose hyperlink references a relationship to a part that is not a hyperlink,
    /// so tests can check that the reader rejects the relationship instead of treating it as a URL.
    /// </summary>
    internal static byte[] CreateDocxWithHyperlinkReferenceToNonHyperlinkPart(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles();
        string id = mainPart.GetIdOfPart(stylesPart);
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Hyperlink(new W.Run(new W.Text(text))) { Id = id })));

        return Flush(document, stream);
    }

    /// <summary>Reads the visible text of every hyperlink in the body.</summary>
    internal static IReadOnlyList<string> ReadDocxHyperlinkTexts(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.Document?.Descendants<W.Hyperlink>() ?? [])
            .Select(link => string.Concat(link.Descendants<W.Text>().Select(text => text.Text)))
            .ToArray();
    }

    /// <summary>Reads the bookmark anchors named by the body's hyperlinks.</summary>
    internal static IReadOnlyList<string> ReadDocxHyperlinkAnchors(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.Document?.Descendants<W.Hyperlink>() ?? [])
            .Select(link => link.Anchor?.Value ?? string.Empty)
            .ToArray();
    }

    /// <summary>
    /// Reads the destination of every body hyperlink in order, resolving its relationship and
    /// returning an empty string for an internal link.
    /// </summary>
    internal static IReadOnlyList<string> ReadDocxHyperlinkDestinations(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        MainDocumentPart? mainPart = document.MainDocumentPart;
        if (mainPart?.Document is null)
        {
            return [];
        }

        Dictionary<string, string> destinations = (mainPart.HyperlinkRelationships ?? [])
            .ToDictionary(relationship => relationship.Id, relationship => relationship.Uri.OriginalString);
        return mainPart.Document.Descendants<W.Hyperlink>()
            .Select(link => link.Id?.Value is { } id && destinations.TryGetValue(id, out string? uri)
                ? uri
                : string.Empty)
            .ToArray();
    }

    /// <summary>Counts the image parts of the main document part.</summary>
    internal static int CountDocxImageParts(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.ImageParts.Count() ?? 0;
    }

    /// <summary>Counts the image parts owned by header parts.</summary>
    internal static int CountDocxHeaderImageParts(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.HeaderParts.Sum(part => part.ImageParts.Count()) ?? 0;
    }

    /// <summary>Counts the image parts owned by footer parts.</summary>
    internal static int CountDocxFooterImageParts(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.FooterParts.Sum(part => part.ImageParts.Count()) ?? 0;
    }

    /// <summary>Reads the bytes of every image part of the main document part.</summary>
    internal static IReadOnlyList<byte[]> ReadDocxImageBytes(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.ImageParts ?? [])
            .Select(part =>
            {
                using Stream data = part.GetStream();
                using MemoryStream copy = new();
                data.CopyTo(copy);
                return copy.ToArray();
            })
            .ToArray();
    }

    /// <summary>Reads the content type of every image part of the main document part.</summary>
    internal static IReadOnlyList<string> ReadDocxImageContentTypes(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.ImageParts ?? []).Select(part => part.ContentType).ToArray();
    }

    /// <summary>Reads the drawing object ids of the body, headers and footers.</summary>
    internal static IReadOnlyList<uint> ReadDocxAllDrawingIds(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        List<uint> ids = [];
        foreach (OpenXmlPartRootElement root in WordRoots(document))
        {
            foreach (WP.DocProperties properties in root.Descendants<WP.DocProperties>())
            {
                if (properties.Id?.Value is { } id)
                {
                    ids.Add(id);
                }
            }
        }

        return ids;
    }

    /// <summary>Reads the inline extent of every inline drawing in the body, headers and footers.</summary>
    internal static IReadOnlyList<(long Width, long Height)> ReadDocxAllInlineExtents(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        List<(long Width, long Height)> extents = [];
        foreach (OpenXmlPartRootElement root in WordRoots(document))
        {
            foreach (WP.Inline inline in root.Descendants<WP.Inline>())
            {
                if (inline.Extent?.Cx?.Value is { } width && inline.Extent.Cy?.Value is { } height)
                {
                    extents.Add((width, height));
                }
            }
        }

        return extents;
    }

    /// <summary>Reads the relationship ids named by the body's hyperlinks.</summary>
    internal static IReadOnlyList<string> ReadDocxHyperlinkReferenceIds(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.Document?.Descendants<W.Hyperlink>() ?? [])
            .Select(link => link.Id?.Value ?? string.Empty)
            .ToArray();
    }

    /// <summary>Reads the destinations of the main document part's external hyperlink relationships.</summary>
    internal static IReadOnlyList<string> ReadDocxExternalHyperlinkUris(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.HyperlinkRelationships ?? [])
            .Where(relationship => relationship.IsExternal)
            .Select(relationship => relationship.Uri.OriginalString)
            .ToArray();
    }

    /// <summary>Reads the destinations of every header part's external hyperlink relationships.</summary>
    internal static IReadOnlyList<string> ReadDocxHeaderExternalHyperlinkUris(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.HeaderParts ?? [])
            .SelectMany(part => part.HyperlinkRelationships)
            .Where(relationship => relationship.IsExternal)
            .Select(relationship => relationship.Uri.OriginalString)
            .ToArray();
    }

    /// <summary>Reads the destinations of every footer part's external hyperlink relationships.</summary>
    internal static IReadOnlyList<string> ReadDocxFooterExternalHyperlinkUris(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.FooterParts ?? [])
            .SelectMany(part => part.HyperlinkRelationships)
            .Where(relationship => relationship.IsExternal)
            .Select(relationship => relationship.Uri.OriginalString)
            .ToArray();
    }

    private static W.Paragraph CreateParagraphWithInlineImage(
        string relationshipId,
        uint id,
        long widthEmus,
        long heightEmus,
        bool includeExtent = true)
    {
        WP.EffectExtent effectExtent = new()
        {
            LeftEdge = 0L,
            TopEdge = 0L,
            RightEdge = 0L,
            BottomEdge = 0L,
        };
        WP.DocProperties docProperties = new() { Id = id, Name = $"Picture {id}" };
        WP.NonVisualGraphicFrameDrawingProperties frameProperties = new(
            new A.GraphicFrameLocks { NoChangeAspect = true });
        PIC.Picture picture = new(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "image.png" },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new PIC.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = widthEmus, Cy = heightEmus }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        A.Graphic graphic = new(
            new A.GraphicData(picture) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });
        WP.Inline inline = new()
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        };
        if (includeExtent)
        {
            inline.AppendChild(new WP.Extent { Cx = widthEmus, Cy = heightEmus });
        }

        inline.Append(effectExtent, docProperties, frameProperties, graphic);

        return new W.Paragraph(new W.Run(new W.Drawing(inline)));
    }

    private static IEnumerable<OpenXmlPartRootElement> WordRoots(WordprocessingDocument document)
    {
        MainDocumentPart? main = document.MainDocumentPart;
        if (main?.Document is { } body)
        {
            yield return body;
        }

        if (main is not null)
        {
            foreach (HeaderPart header in main.HeaderParts)
            {
                if (header.Header is { } headerRoot)
                {
                    yield return headerRoot;
                }
            }

            foreach (FooterPart footer in main.FooterParts)
            {
                if (footer.Footer is { } footerRoot)
                {
                    yield return footerRoot;
                }
            }
        }
    }

    /// <summary>Counts the header parts of a .docx package, whether or not the body references them.</summary>
    internal static int CountDocxHeaderParts(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.HeaderParts.Count() ?? 0;
    }

    /// <summary>Counts the footer parts of a .docx package, whether or not the body references them.</summary>
    internal static int CountDocxFooterParts(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.FooterParts.Count() ?? 0;
    }

    /// <summary>Reads whether any header part of a .docx package holds a content control this version does not model.</summary>
    internal static bool ReadDocxHeaderHasContentControl(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.HeaderParts
            .Any(part => part.Header?.GetFirstChild<W.SdtBlock>() is not null) ?? false;
    }

    /// <summary>
    /// Reads the text of every header part of a .docx package, as the plain concatenation of the
    /// paragraphs of each part.
    /// </summary>
    internal static IReadOnlyList<string> ReadDocxHeaderTexts(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.HeaderParts ?? [])
            .Select(part => string.Join("\n", part.Header?.Elements<W.Paragraph>().Select(p => p.InnerText) ?? []))
            .ToArray();
    }

    /// <summary>
    /// Reads the text of every footer part of a .docx package, as the plain concatenation of the
    /// paragraphs of each part.
    /// </summary>
    internal static IReadOnlyList<string> ReadDocxFooterTexts(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return (document.MainDocumentPart?.FooterParts ?? [])
            .Select(part => string.Join("\n", part.Footer?.Elements<W.Paragraph>().Select(p => p.InnerText) ?? []))
            .ToArray();
    }

    /// <summary>Reads the fields of the footer parts with the Open XML SDK.</summary>
    internal static (IReadOnlyList<string> Instructions, IReadOnlyList<string> FieldCharTypes) ReadDocxFooterFields(
        byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);

        List<string> instructions = [];
        List<string> types = [];
        foreach (FooterPart part in document.MainDocumentPart?.FooterParts ?? [])
        {
            instructions.AddRange(part.Footer?.Descendants<W.FieldCode>().Select(code => code.Text) ?? []);
            types.AddRange(part.Footer?.Descendants<W.FieldChar>()
                .Select(fieldChar => fieldChar.FieldCharType?.InnerText ?? string.Empty) ?? []);
        }

        return (instructions, types);
    }

    /// <summary>
    /// Builds a .docx with a paragraph, a body-level content control (<c>w:sdt</c>, which this version
    /// does not model) and another paragraph.
    /// </summary>
    internal static byte[] CreateDocxWithUnmodelledBodyContent(string text)
        => CreateDocxWithBody(
            new W.Paragraph(new W.Run(new W.Text(text))),
            new W.SdtBlock(new W.SdtContentBlock(new W.Paragraph(new W.Run(new W.Text("inside"))))));

    /// <summary>
    /// Builds a .docx with one table whose cell holds a paragraph and a nested table, which this
    /// version preserves but does not model.
    /// </summary>
    internal static byte[] CreateDocxWithNestedTable(string outerText, string innerText)
        => CreateDocxWithBody(new W.Table(
            new W.TableProperties(),
            new W.TableGrid(new W.GridColumn { Width = "2400" }),
            new W.TableRow(new W.TableCell(
                new W.Paragraph(new W.Run(new W.Text(outerText))),
                new W.Table(
                    new W.TableProperties(),
                    new W.TableGrid(new W.GridColumn { Width = "1200" }),
                    new W.TableRow(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(innerText))))))))));

    /// <summary>
    /// Builds a .docx with a single table whose markup carries properties this version does not model
    /// (table width, cell width, shading and a horizontal merge), so tests can check that editing the
    /// content leaves them untouched.
    /// </summary>
    internal static byte[] CreateDocxWithFormattedTable(string cellText)
        => CreateDocxWithBody(new W.Table(
            new W.TableProperties(new W.TableWidth { Width = "0", Type = W.TableWidthUnitValues.Auto }),
            new W.TableGrid(new W.GridColumn { Width = "2400" }, new W.GridColumn { Width = "2400" }),
            new W.TableRow(
                new W.TableCell(
                    new W.TableCellProperties(
                        new W.TableCellWidth { Width = "2400", Type = W.TableWidthUnitValues.Dxa },
                        new W.GridSpan { Val = 2 },
                        new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "D9E2F3" }),
                    new W.Paragraph(new W.Run(new W.Text(cellText)))))));

    /// <summary>
    /// Builds a .docx with one table whose <c>w:tblPr</c> is exactly the given markup, so tests can
    /// place values OfficeEngine models and markup it preserves without going through the code under
    /// test.
    /// </summary>
    internal static byte[] CreateDocxWithTableProperties(string cellText, W.TableProperties properties)
        => CreateDocxWithBody(new W.Table(
            properties,
            new W.TableGrid(new W.GridColumn { Width = "2400" }),
            new W.TableRow(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(cellText)))))));

    /// <summary>Builds a .docx with one table whose only cell carries the given properties markup.</summary>
    internal static byte[] CreateDocxWithCellProperties(string cellText, W.TableCellProperties properties)
        => CreateDocxWithBody(new W.Table(
            new W.TableProperties(),
            new W.TableGrid(new W.GridColumn { Width = "2400" }),
            new W.TableRow(new W.TableCell(properties, new W.Paragraph(new W.Run(new W.Text(cellText)))))));

    /// <summary>
    /// Builds a .docx with one two-column, three-row table whose first column is vertically merged: the
    /// top cell restarts the merge, the middle one omits <c>w:val</c> and the bottom one states it, so
    /// tests can read markup the code under test did not write.
    /// </summary>
    internal static byte[] CreateDocxWithVerticallyMergedTable(string topText, string middleText, string bottomText)
        => CreateDocxWithBody(new W.Table(
            new W.TableProperties(),
            new W.TableGrid(new W.GridColumn { Width = "2400" }, new W.GridColumn { Width = "2400" }),
            new W.TableRow(
                new W.TableCell(
                    new W.TableCellProperties(new W.VerticalMerge { Val = W.MergedCellValues.Restart }),
                    new W.Paragraph(new W.Run(new W.Text(topText)))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text("right top"))))),
            new W.TableRow(
                new W.TableCell(
                    new W.TableCellProperties(new W.VerticalMerge()),
                    new W.Paragraph(new W.Run(new W.Text(middleText)))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text("right middle"))))),
            new W.TableRow(
                new W.TableCell(
                    new W.TableCellProperties(new W.VerticalMerge { Val = W.MergedCellValues.Continue }),
                    new W.Paragraph(new W.Run(new W.Text(bottomText)))),
                new W.TableCell(new W.Paragraph(new W.Run(new W.Text("right bottom")))))));

    /// <summary>Reads the <c>w:tblPr</c> of the first table of a .docx package.</summary>
    internal static W.TableProperties? ReadDocxFirstTableProperties(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.Document?.Body?.GetFirstChild<W.Table>()?.GetFirstChild<W.TableProperties>();
    }

    /// <summary>Reads the <c>w:tcPr</c> of the first cell of the first row of the first table.</summary>
    internal static W.TableCellProperties? ReadDocxFirstTableCellProperties(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.Document?.Body?.GetFirstChild<W.Table>()
            ?.GetFirstChild<W.TableRow>()?.GetFirstChild<W.TableCell>()?.GetFirstChild<W.TableCellProperties>();
    }

    /// <summary>Reads the local names of the body's direct children, in document order.</summary>
    internal static IReadOnlyList<string> ReadDocxBodyChildLocalNames(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        W.Body? body = document.MainDocumentPart?.Document?.Body;
        return body is null ? [] : body.ChildElements.Select(child => child.LocalName).ToArray();
    }

    /// <summary>
    /// Reads the section properties of a .docx package in document order: the properties stored in
    /// the paragraph that ends each earlier section, followed by the final <c>w:sectPr</c> of the body.
    /// </summary>
    internal static IReadOnlyList<W.SectionProperties> ReadDocxSections(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        W.Body? body = document.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            return [];
        }

        List<W.SectionProperties> sections = [];
        foreach (OpenXmlElement child in body.ChildElements)
        {
            if (child is W.Paragraph paragraph && paragraph.ParagraphProperties?.SectionProperties is { } stored)
            {
                sections.Add(stored);
            }
            else if (child is W.SectionProperties final)
            {
                sections.Add(final);
            }
        }

        return sections;
    }

    /// <summary>Reads the final <c>w:sectPr</c> of the body of a .docx package.</summary>
    internal static W.SectionProperties? ReadDocxFinalSectionProperties(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.Document?.Body?.Elements<W.SectionProperties>().LastOrDefault();
    }

    /// <summary>Reads the direct tables of the body of a .docx package with the Open XML SDK.</summary>
    internal static IReadOnlyList<W.Table> ReadDocxTables(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.Document?.Body?.Elements<W.Table>().ToList() ?? [];
    }

    /// <summary>
    /// Builds a .docx whose single paragraph defines the given tab stops, so tests can exercise the
    /// values OfficeEngine models and the markup it does not.
    /// </summary>
    internal static byte[] CreateDocxWithTabStops(string text, params W.TabStop[] stops)
        => CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(new W.Tabs(stops)));

    /// <summary>
    /// Builds a .docx whose single paragraph stores a tab stop with the given raw attributes, so
    /// tests can cover a value OfficeEngine models, one it does not model, and one that cannot be
    /// parsed. A <see langword="null"/> attribute is omitted.
    /// </summary>
    internal static byte[] CreateDocxWithStoredTabStop(string text, string? alignment, string? position, string? leader)
    {
        W.TabStop stop = new();
        if (alignment is not null)
        {
            stop.SetAttribute(new OpenXmlAttribute("w", "val", WordprocessingNamespace, alignment));
        }

        if (position is not null)
        {
            stop.SetAttribute(new OpenXmlAttribute("w", "pos", WordprocessingNamespace, position));
        }

        if (leader is not null)
        {
            stop.SetAttribute(new OpenXmlAttribute("w", "leader", WordprocessingNamespace, leader));
        }

        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(new W.Tabs(stop)));
    }

    /// <summary>
    /// Builds a .docx whose first run stores a line break with the given raw <c>w:type</c>, so tests
    /// can cover a value OfficeEngine models, a value the schema spells out explicitly and one that
    /// cannot be parsed. A <see langword="null"/> type omits the attribute.
    /// </summary>
    internal static byte[] CreateDocxWithStoredBreak(string text, string? type)
    {
        W.Break lineBreak = new();
        if (type is not null)
        {
            lineBreak.SetAttribute(new OpenXmlAttribute("w", "type", WordprocessingNamespace, type));
        }

        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text(text), lineBreak))));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose first run mixes modelled content (text, a tab) with markup OfficeEngine
    /// preserves but does not model (a carriage return and a rendered page break).
    /// </summary>
    internal static byte[] CreateDocxWithMixedRunContent(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(
            new W.Text(text),
            new W.CarriageReturn(),
            new W.LastRenderedPageBreak(),
            new W.TabChar(),
            new W.Text("tail")))));

        return Flush(document, stream);
    }

    /// <summary>Reads the first paragraph of a .docx package with the Open XML SDK.</summary>
    internal static W.Paragraph? ReadDocxFirstParagraph(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.Document?.Body?.Descendants<W.Paragraph>().FirstOrDefault();
    }

    /// <summary>Reads the first run of the first paragraph of a .docx package with the Open XML SDK.</summary>
    internal static W.Run? ReadDocxFirstRun(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.Document?.Body?.Descendants<W.Paragraph>().FirstOrDefault()
            ?.Elements<W.Run>().FirstOrDefault();
    }

    /// <summary>Builds a .xlsx that uses a shared string table and a merged cell range.</summary>
    internal static byte[] CreateXlsxWithSharedStringsAndMergedCells(string firstText, string secondText)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        SharedStringTablePart sharedStrings = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStrings.SharedStringTable = new S.SharedStringTable(
            new S.SharedStringItem(new S.Text(firstText)),
            new S.SharedStringItem(new S.Text(secondText)));

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(
            new S.SheetData(
                new S.Row(
                    new S.Cell { CellReference = "A1", DataType = S.CellValues.SharedString, CellValue = new S.CellValue("0") },
                    new S.Cell { CellReference = "B1", DataType = S.CellValues.SharedString, CellValue = new S.CellValue("1") })
                { RowIndex = 1U }),
            new S.MergeCells(new S.MergeCell { Reference = "A1:B1" }));

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>Builds an .xlsx whose single cell contains a formula.</summary>
    internal static byte[] CreateXlsxWithFormula(
        string formula,
        string cachedValue,
        bool includeDefinedName = false,
        bool includeTrailingElement = false)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData(new S.Row(
            new S.Cell { CellReference = "A1", CellFormula = new S.CellFormula(formula), CellValue = new S.CellValue(cachedValue) })
            { RowIndex = 1U }));

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));
        if (includeDefinedName)
        {
            workbookPart.Workbook.AppendChild(new S.DefinedNames(
                new S.DefinedName { Name = "Datos", Text = "Hoja1!$A$1" }));
        }

        if (includeTrailingElement)
        {
            workbookPart.Workbook.AppendChild(new S.FileRecoveryProperties());
        }

        return Flush(document, stream);
    }

    /// <summary>Builds an intentionally malformed .xlsx whose row omits the required row index.</summary>
    internal static byte[] CreateXlsxWithRowWithoutIndex(string text)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData(new S.Row(
            new S.Cell
            {
                CellReference = "A1",
                DataType = S.CellValues.InlineString,
                InlineString = new S.InlineString(new S.Text(text)),
            })));

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds an .xlsx with one styled cell. The stylesheet always contains a default style and a
    /// second style that applies the requested number format; <paramref name="includeOtherFormatting"/>
    /// also adds bold, yellow fill, a thick bottom border and centered alignment to that second style.
    /// <paramref name="includeUnmodelledFormatting"/> adds a strike-through, a diagonal border and an
    /// indent, which this version does not model, to that second style. By default the cell uses the
    /// second style; <paramref name="useDefaultStyle"/> makes the cell use the default style and puts
    /// the requested number format there instead.
    /// </summary>
    internal static byte[] CreateXlsxWithFormattedCell(
        string cellReference,
        string cellValue,
        uint numberFormatId,
        string? customFormatCode = null,
        bool date1904 = false,
        bool includeOtherFormatting = false,
        bool useDefaultStyle = false,
        bool includeUnmodelledFormatting = false)
    {
        using MemoryStream stream = new();
        SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();
        if (date1904)
        {
            workbookPart.Workbook.AppendChild(new S.WorkbookProperties { Date1904 = true });
        }

        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        S.Stylesheet stylesheet = new();
        if (customFormatCode is not null)
        {
            stylesheet.AppendChild(new S.NumberingFormats(
                new S.NumberingFormat
                {
                    NumberFormatId = numberFormatId,
                    FormatCode = customFormatCode,
                })
            { Count = 1U });
        }

        S.Font otherFont = includeUnmodelledFormatting
            ? new S.Font(new S.Bold(), new S.Strike())
            : new S.Font(new S.Bold());
        stylesheet.AppendChild(new S.Fonts(
            new S.Font(),
            otherFont)
        { Count = 2U });
        stylesheet.AppendChild(new S.Fills(
            new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None }),
            new S.Fill(new S.PatternFill
            {
                PatternType = S.PatternValues.Solid,
                ForegroundColor = new S.ForegroundColor { Rgb = "FFFFFF00" },
                BackgroundColor = new S.BackgroundColor { Indexed = 64U },
            }))
        { Count = 2U });
        S.Border otherBorder = new(new S.BottomBorder { Style = S.BorderStyleValues.Thick });
        if (includeUnmodelledFormatting)
        {
            otherBorder.AppendChild(new S.DiagonalBorder { Style = S.BorderStyleValues.Dashed });
        }

        stylesheet.AppendChild(new S.Borders(
            new S.Border(),
            otherBorder)
        { Count = 2U });
        stylesheet.AppendChild(new S.CellStyleFormats(
            new S.CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U })
        { Count = 1U });

        S.CellFormat defaultFormat = new()
        {
            NumberFormatId = useDefaultStyle ? numberFormatId : 0U,
            FontId = 0U,
            FillId = 0U,
            BorderId = 0U,
            FormatId = 0U,
        };
        if (useDefaultStyle)
        {
            defaultFormat.ApplyNumberFormat = true;
        }

        S.CellFormat formatted = new()
        {
            NumberFormatId = numberFormatId,
            FontId = includeOtherFormatting ? 1U : 0U,
            FillId = includeOtherFormatting ? 1U : 0U,
            BorderId = includeOtherFormatting ? 1U : 0U,
            FormatId = 0U,
            ApplyNumberFormat = true,
        };
        if (includeOtherFormatting)
        {
            formatted.ApplyFont = true;
            formatted.ApplyFill = true;
            formatted.ApplyBorder = true;
            formatted.ApplyAlignment = true;
            S.Alignment alignment = new() { Horizontal = S.HorizontalAlignmentValues.Center };
            if (includeUnmodelledFormatting)
            {
                alignment.Indent = 2U;
            }

            formatted.Alignment = alignment;
        }

        stylesheet.AppendChild(new S.CellFormats(
            defaultFormat,
            formatted)
        { Count = 2U });
        stylesheet.AppendChild(new S.CellStyles(
            new S.CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U })
        { Count = 1U });
        stylesPart.Stylesheet = stylesheet;

        S.Cell cell = new()
        {
            CellReference = cellReference,
            CellValue = new S.CellValue(cellValue),
        };
        if (!useDefaultStyle)
        {
            cell.StyleIndex = 1U;
        }

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData(new S.Row(cell)
        { RowIndex = 1U }));

        workbookPart.Workbook.AppendChild(new S.Sheets(
            new S.Sheet { Name = "Hoja1", SheetId = 1U, Id = workbookPart.GetIdOfPart(worksheetPart) }));

        return Flush(document, stream);
    }

    /// <summary>
    /// Reads the style of a cell from an .xlsx package with the Open XML SDK, including the custom
    /// format code when the number format is not built in.
    /// </summary>
    internal static (uint StyleIndex, uint NumberFormatId, uint FontId, uint FillId, uint BorderId, string? FormatCode)
        ReadXlsxCellStyle(byte[] package, string sheetName, string cellReference)
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        S.Cell cell = worksheetPart.Worksheet!.Descendants<S.Cell>()
            .First(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.Ordinal));

        uint styleIndex = cell.StyleIndex?.Value ?? 0U;
        S.Stylesheet stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet
            ?? throw new Xunit.Sdk.XunitException("The package has no stylesheet.");
        S.CellFormat format = stylesheet.CellFormats!.Elements<S.CellFormat>().ElementAt((int)styleIndex);
        uint numberFormatId = format.NumberFormatId?.Value ?? 0U;
        string? formatCode = stylesheet.NumberingFormats?.Elements<S.NumberingFormat>()
            .FirstOrDefault(existing => existing.NumberFormatId?.Value == numberFormatId)
            ?.FormatCode?.Value;

        return (
            styleIndex,
            numberFormatId,
            format.FontId?.Value ?? 0U,
            format.FillId?.Value ?? 0U,
            format.BorderId?.Value ?? 0U,
            formatCode);
    }

    /// <summary>
    /// Reads the cell format and the font, fill and border it references from an .xlsx package with
    /// the Open XML SDK, so style assertions do not depend on the code under test.
    /// </summary>
    internal static (uint StyleIndex, S.CellFormat Format, S.Font Font, S.Fill Fill, S.Border Border)
        ReadXlsxCellStyleParts(byte[] package, string sheetName, string cellReference)
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        WorkbookPart workbookPart = document.WorkbookPart
            ?? throw new Xunit.Sdk.XunitException("The package has no workbook part.");
        S.Sheet sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        S.Cell cell = worksheetPart.Worksheet!.Descendants<S.Cell>()
            .First(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.Ordinal));

        uint styleIndex = cell.StyleIndex?.Value ?? 0U;
        S.Stylesheet stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet
            ?? throw new Xunit.Sdk.XunitException("The package has no stylesheet.");
        S.CellFormat format = stylesheet.CellFormats!.Elements<S.CellFormat>().ElementAt((int)styleIndex);
        S.Font font = stylesheet.Fonts!.Elements<S.Font>()
            .ElementAt((int)(format.FontId?.Value ?? 0U));
        S.Fill fill = stylesheet.Fills!.Elements<S.Fill>()
            .ElementAt((int)(format.FillId?.Value ?? 0U));
        S.Border border = stylesheet.Borders!.Elements<S.Border>()
            .ElementAt((int)(format.BorderId?.Value ?? 0U));

        return (styleIndex, format, font, fill, border);
    }

    /// <summary>Counts the font, fill, border and cell-format entries in an .xlsx stylesheet.</summary>
    internal static (int Fonts, int Fills, int Borders, int CellFormats) ReadXlsxStyleCounts(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        S.Stylesheet stylesheet = document.WorkbookPart?.WorkbookStylesPart?.Stylesheet
            ?? throw new Xunit.Sdk.XunitException("The package has no stylesheet.");
        return (
            stylesheet.Fonts?.Elements<S.Font>().Count() ?? 0,
            stylesheet.Fills?.Elements<S.Fill>().Count() ?? 0,
            stylesheet.Borders?.Elements<S.Border>().Count() ?? 0,
            stylesheet.CellFormats?.Elements<S.CellFormat>().Count() ?? 0);
    }

    /// <summary>
    /// Builds a .docx whose style part stores the given style definitions and whose body holds a
    /// single paragraph with the given text, so tests can exercise the style catalogue without going
    /// through the code under test.
    /// </summary>
    internal static byte[] CreateDocxWithStyles(string text, params W.Style[] styles)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new W.Styles(styles);
        mainPart.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text(text)))));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx with a style part that has no <c>w:styles</c> root, so tests can check that a
    /// document without a readable catalogue reports an empty one.
    /// </summary>
    internal static byte[] CreateDocxWithEmptyStylePart(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.AddNewPart<StyleDefinitionsPart>();
        mainPart.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text(text)))));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose single paragraph and run reference the given style ids. A
    /// <see langword="null"/> id omits the reference, so the helper can produce a paragraph without a
    /// style, a run without one, or both.
    /// </summary>
    internal static byte[] CreateDocxWithStyleReference(string text, string? paragraphStyleId, string? runStyleId)
    {
        W.Paragraph paragraph = new();
        if (paragraphStyleId is not null)
        {
            paragraph.AppendChild(new W.ParagraphProperties(new W.ParagraphStyleId { Val = paragraphStyleId }));
        }

        W.Run run = new();
        if (runStyleId is not null)
        {
            run.AppendChild(new W.RunProperties(new W.RunStyle { Val = runStyleId }));
        }

        run.AppendChild(new W.Text(text));
        paragraph.AppendChild(run);

        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(paragraph));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose single paragraph stores a <c>w:pStyle</c> with the given raw value, so
    /// tests can cover a reference with and without its required attribute.
    /// </summary>
    internal static byte[] CreateDocxWithStoredParagraphStyleReference(string text, string? value)
    {
        W.ParagraphStyleId style = new();
        if (value is not null)
        {
            style.SetAttribute(new OpenXmlAttribute("w", "val", WordprocessingNamespace, value));
        }

        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(style));
    }

    /// <summary>
    /// Builds a .docx whose single run stores a <c>w:rStyle</c> with the given raw value, so tests can
    /// cover a reference with and without its required attribute.
    /// </summary>
    internal static byte[] CreateDocxWithStoredRunStyleReference(string text, string? value)
    {
        W.RunStyle style = new();
        if (value is not null)
        {
            style.SetAttribute(new OpenXmlAttribute("w", "val", WordprocessingNamespace, value));
        }

        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(
            new W.RunProperties(style),
            new W.Text(text)))));

        return Flush(document, stream);
    }

    /// <summary>Reads the style catalogue of a .docx package with the Open XML SDK.</summary>
    internal static W.Styles? ReadDocxStyles(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.StyleDefinitionsPart?.Styles;
    }

    /// <summary>Reads the raw <c>word/styles.xml</c> entry of a package as text.</summary>
    internal static string ReadDocxStylesXml(byte[] package)
        => TestPackage.ReadEntryText(package, "word/styles.xml");

    /// <summary>
    /// Builds a .docx with a numbering part (<c>word/numbering.xml</c>) whose root holds the given
    /// children (<c>w:abstractNum</c>/<c>w:num</c>) and whose single paragraph optionally carries the
    /// given numbering properties. Passing <see langword="null"/> for the properties leaves the
    /// paragraph without a list reference.
    /// </summary>
    internal static byte[] CreateDocxWithLists(string text, W.NumberingProperties? paragraphNumbering, params OpenXmlElement[] numberingChildren)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new W.Numbering(numberingChildren);

        W.Paragraph paragraph = new();
        if (paragraphNumbering is not null)
        {
            paragraph.AppendChild(new W.ParagraphProperties(paragraphNumbering));
        }

        paragraph.AppendChild(new W.Run(new W.Text(text)));
        mainPart.Document = new W.Document(new W.Body(paragraph));

        return Flush(document, stream);
    }

    /// <summary>
    /// Builds a .docx whose single paragraph carries the <c>w:numPr</c> the configure action builds,
    /// so tests can place exactly the numbering reference they need (including one with a missing or
    /// malformed attribute) without going through the code under test.
    /// </summary>
    internal static byte[] CreateDocxWithNumberingProperties(string text, Action<W.NumberingProperties> configure)
    {
        W.NumberingProperties numbering = new();
        configure(numbering);
        return CreateDocxWithParagraphProperties(text, new W.ParagraphProperties(numbering));
    }

    /// <summary>
    /// Sets a raw value on an attribute of an element, or leaves it unset when the value is
    /// <see langword="null"/>, so a fixture can be built with a value the typed API cannot express.
    /// </summary>
    internal static T WithAttribute<T>(T element, string localName, string? value)
        where T : OpenXmlElement
    {
        if (value is not null)
        {
            element.SetAttribute(new OpenXmlAttribute("w", localName, WordprocessingNamespace, value));
        }

        return element;
    }

    /// <summary>
    /// Builds an abstract numbering definition (<c>w:abstractNum</c>) with a single, schema-ordered
    /// level that uses the given number format, so tests can assemble a numbering part without going
    /// through the code under test.
    /// </summary>
    internal static W.AbstractNum AbstractNumbering(int abstractNumberId, W.NumberFormatValues format, int level = 0)
    {
        W.Level definition = new()
        {
            LevelIndex = level,
            StartNumberingValue = new W.StartNumberingValue { Val = 1 },
            NumberingFormat = new W.NumberingFormat { Val = format },
            LevelText = new W.LevelText { Val = format == W.NumberFormatValues.Bullet ? "\uF0B7" : "%1." },
            LevelJustification = new W.LevelJustification { Val = W.LevelJustificationValues.Left },
        };

        return new W.AbstractNum(definition) { AbstractNumberId = abstractNumberId };
    }

    /// <summary>Builds a list definition (<c>w:num</c>) that references the given abstract numbering definition.</summary>
    internal static W.NumberingInstance ListInstance(int listId, int abstractNumberId)
        => new(new W.AbstractNumId { Val = abstractNumberId }) { NumberID = listId };

    /// <summary>
    /// Builds a .docx with a numbering part that has no <c>w:numbering</c> root, so tests can check
    /// that a document without a readable catalogue reports an empty one.
    /// </summary>
    internal static byte[] CreateDocxWithEmptyNumberingPart(string text)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.AddNewPart<NumberingDefinitionsPart>();
        mainPart.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text(text)))));

        return Flush(document, stream);
    }

    /// <summary>Reads the numbering root of a .docx package with the Open XML SDK.</summary>
    internal static W.Numbering? ReadDocxNumbering(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
    }

    /// <summary>Reads the raw <c>word/numbering.xml</c> entry of a package as text.</summary>
    internal static string ReadDocxNumberingXml(byte[] package)
        => TestPackage.ReadEntryText(package, "word/numbering.xml");

    /// <summary>
    /// Builds a .docx with modelled core properties and with extended and custom properties that
    /// OfficeEngine does not model, so tests can check both the readable values and preservation.
    /// </summary>
    internal static byte[] CreateDocxWithDocumentProperties(
        string title,
        string creator,
        DateTime created,
        string company,
        string customValue)
    {
        using MemoryStream stream = new();
        WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(
            new W.Paragraph(new W.Run(new W.Text("cuerpo")))));

        CoreFilePropertiesPart corePart = document.AddNewPart<CoreFilePropertiesPart>();
        using (Stream coreStream = corePart.GetStream(FileMode.Create, FileAccess.Write))
        {
            byte[] root = Encoding.UTF8.GetBytes(EmptyCorePropertiesXml);
            coreStream.Write(root, 0, root.Length);
        }

        corePart.CoreFileProperties.Title = title;
        corePart.CoreFileProperties.Creator = creator;
        corePart.CoreFileProperties.Created = created;

        ExtendedFilePropertiesPart extendedPart = document.AddNewPart<ExtendedFilePropertiesPart>();
        extendedPart.Properties = new EP.Properties(
            new EP.Company(company),
            new EP.Manager("Manager"),
            new EP.Application("OfficeEngine"));

        CustomFilePropertiesPart customPart = document.AddNewPart<CustomFilePropertiesPart>();
        customPart.Properties = new CP.Properties(
            new CP.CustomDocumentProperty(new VT.VTLPWSTR(customValue))
            {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = 2,
                Name = "CustomName",
            });

        return Flush(document, stream);
    }

    /// <summary>
    /// Rebuilds a .docx with a malformed <c>docProps/core.xml</c> date, so tests can check the typed
    /// failure instead of an SDK exception.
    /// </summary>
    internal static byte[] CreateDocxWithMalformedCorePropertyDate(string title)
    {
        byte[] valid = CreateDocxWithDocumentProperties(
            title,
            "Autor",
            new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc),
            "Acme",
            "valor");

        string malformed =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
            "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
            "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" " +
            "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
            "<dc:title>" + title + "</dc:title>" +
            "<dcterms:created xsi:type=\"dcterms:W3CDTF\">not-a-date</dcterms:created>" +
            "</cp:coreProperties>";

        return TestPackage.WithEntryText(valid, "docProps/core.xml", malformed);
    }

    /// <summary>Reads the raw <c>docProps/core.xml</c> entry.</summary>
    internal static string ReadDocxCorePropertiesXml(byte[] package)
        => TestPackage.ReadEntryText(package, "docProps/core.xml");

    /// <summary>
    /// Parses <c>docProps/core.xml</c> with an XML parser independent of OfficeEngine and returns
    /// each property value keyed by its local name.
    /// </summary>
    internal static IReadOnlyDictionary<string, string?> ReadDocxCoreProperties(byte[] package)
    {
        XDocument core = XDocument.Parse(ReadDocxCorePropertiesXml(package));
        return core.Root!.Elements().ToDictionary(
            element => element.Name.LocalName,
            element => (string?)element.Value,
            StringComparer.Ordinal);
    }

    /// <summary>Reads a custom property value with the Open XML SDK.</summary>
    internal static string? ReadDocxCustomPropertyValue(byte[] package, string name)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        CP.CustomDocumentProperty? property = document.CustomFilePropertiesPart?.Properties?
            .Elements<CP.CustomDocumentProperty>()
            .FirstOrDefault(candidate => candidate.Name?.Value == name);
        return property?.InnerText;
    }

    /// <summary>Reads the extended company property with the Open XML SDK.</summary>
    internal static string? ReadDocxCompany(byte[] package)
    {
        using MemoryStream stream = new(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        return document.ExtendedFilePropertiesPart?.Properties?.Company?.Text;
    }

    private static byte[] Flush(OpenXmlPackage package, MemoryStream stream)
    {
        package.Save();
        byte[] bytes = stream.ToArray();
        package.Dispose();
        return bytes;
    }
}
