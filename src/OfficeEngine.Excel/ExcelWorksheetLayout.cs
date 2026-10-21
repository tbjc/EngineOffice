using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>Keeps the children of a worksheet part in the order that the schema requires.</summary>
internal static class ExcelWorksheetLayout
{
    private static readonly string[] ElementOrder =
    [
        "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr",
        "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
        "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
        "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks",
        "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
        "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls",
        "webPublishItems", "tableParts", "extLst",
    ];

    internal static void InsertInOrder(S.Worksheet worksheet, OpenXmlElement element)
    {
        int order = Array.IndexOf(ElementOrder, element.LocalName);
        OpenXmlElement? next = worksheet.ChildElements
            .FirstOrDefault(existing => Array.IndexOf(ElementOrder, existing.LocalName) > order);
        if (next is null)
        {
            worksheet.AppendChild(element);
        }
        else
        {
            worksheet.InsertBefore(element, next);
        }
    }
}
