using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>Keeps the children of the workbook part in the order that the schema requires.</summary>
internal static class ExcelWorkbookLayout
{
    private static readonly string[] ElementOrder =
    [
        "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets",
        "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize",
        "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
        "fileRecoveryPr", "webPublishObjects", "extLst",
    ];

    internal static void InsertInOrder(S.Workbook root, OpenXmlElement element)
    {
        int order = Array.IndexOf(ElementOrder, element.LocalName);
        OpenXmlElement? next = root.ChildElements
            .FirstOrDefault(existing => Array.IndexOf(ElementOrder, existing.LocalName) > order);
        if (next is null)
        {
            root.AppendChild(element);
        }
        else
        {
            root.InsertBefore(element, next);
        }
    }
}
