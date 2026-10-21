namespace OfficeEngine.Legacy;

/// <summary>The legacy formats that can be converted to their modern counterpart.</summary>
public enum LegacyDocumentFormat
{
    /// <summary>A Word 97-2003 document (`.doc`), converted to a Word document (`.docx`).</summary>
    Word97To2003 = 0,

    /// <summary>An Excel 97-2003 workbook (`.xls`), converted to a workbook (`.xlsx`).</summary>
    Excel97To2003 = 1,
}
