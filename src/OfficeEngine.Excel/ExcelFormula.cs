using System.Xml;
using OfficeEngine.Formulas;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

internal static class ExcelFormula
{
    private const int MaxLength = 8192;

    /// <summary>Normalizes a public formula, which may include a leading equals sign.</summary>
    internal static string Normalize(string value)
    {
        ArgumentException.ThrowIfNullOrEmpty(value);
        string formula = value[0] == '=' ? value[1..] : value;
        if (formula.Length == 0 || string.IsNullOrWhiteSpace(formula))
        {
            throw new ArgumentException("The formula must contain an expression.", nameof(value));
        }

        if (formula.Length > MaxLength)
        {
            throw new ArgumentException(
                $"The formula must be at most {MaxLength} characters long.",
                nameof(value));
        }

        try
        {
            _ = XmlConvert.VerifyXmlChars(formula);
        }
        catch (XmlException ex)
        {
            throw new ArgumentException(
                "The formula contains characters that are not valid in XML.",
                nameof(value),
                ex);
        }

        // Excel stores the names of the functions in English, so a formula written with the Spanish
        // names is translated before it is stored (ADR 0043).
        return FormulaNames.Canonicalize(formula);
    }

    /// <summary>Returns the public form of a stored formula, or null when it has no text.</summary>
    internal static string? ToPublic(string? formula)
    {
        if (string.IsNullOrEmpty(formula))
        {
            return null;
        }

        return formula[0] == '=' ? formula : "=" + formula;
    }

    /// <summary>
    /// Marks the workbook for a full recalculation when it is opened by an engine that can calculate
    /// formulas. This library never calculates them.
    /// </summary>
    internal static void RequestRecalculation(ExcelWorkbook workbook)
    {
        S.Workbook root = workbook.WorkbookPart.Workbook
            ?? throw new InvalidOfficeFileException("The package does not contain a readable workbook part.");

        S.CalculationProperties? properties = root.GetFirstChild<S.CalculationProperties>();
        if (properties is null)
        {
            properties = new S.CalculationProperties();
            ExcelWorkbookLayout.InsertInOrder(root, properties);
        }

        properties.FullCalculationOnLoad = true;
    }
}
