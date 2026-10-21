using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>Maps data validation rules to and from their SpreadsheetML representation.</summary>
internal static class ExcelDataValidations
{
    /// <summary>Validates an A1-style range and returns its normalized text.</summary>
    internal static string NormalizeRange(string rangeReference)
        => CellReference.NormalizeRangeOrCell(rangeReference, nameof(rangeReference));

    /// <summary>Builds the SpreadsheetML element of a rule.</summary>
    internal static S.DataValidation ToElement(ExcelDataValidation validation)
    {
        S.DataValidation element = new()
        {
            Type = ToValues(validation.Type),
            AllowBlank = validation.AllowBlank,
            ShowInputMessage = true,
            ShowErrorMessage = true,
            SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue(validation.Range) }),
        };

        if (validation.Type is ExcelDataValidationType.WholeNumber or ExcelDataValidationType.DecimalNumber)
        {
            if (validation.Minimum is double minimum && validation.Maximum is double maximum)
            {
                element.Operator = S.DataValidationOperatorValues.Between;
                element.Formula1 = new S.Formula1(Format(minimum));
                element.Formula2 = new S.Formula2(Format(maximum));
            }
            else if (validation.Minimum is double lower)
            {
                element.Operator = S.DataValidationOperatorValues.GreaterThanOrEqual;
                element.Formula1 = new S.Formula1(Format(lower));
            }
            else if (validation.Maximum is double upper)
            {
                element.Operator = S.DataValidationOperatorValues.LessThanOrEqual;
                element.Formula1 = new S.Formula1(Format(upper));
            }
        }
        else
        {
            if (validation.Formula1 is not null)
            {
                element.Formula1 = new S.Formula1(validation.Formula1);
            }

            if (validation.Formula2 is not null)
            {
                element.Formula2 = new S.Formula2(validation.Formula2);
            }
        }

        if (validation.ErrorTitle is not null)
        {
            element.ErrorTitle = validation.ErrorTitle;
        }

        if (validation.Error is not null)
        {
            element.Error = validation.Error;
        }

        if (validation.PromptTitle is not null)
        {
            element.PromptTitle = validation.PromptTitle;
        }

        if (validation.Prompt is not null)
        {
            element.Prompt = validation.Prompt;
        }

        return element;
    }

    /// <summary>Reads a rule from its SpreadsheetML element.</summary>
    internal static ExcelDataValidation FromElement(S.DataValidation element)
    {
        string range = FirstArea(element.SequenceOfReferences?.InnerText);
        ExcelDataValidationType type = ToType(element.Type?.Value);
        string? formula1 = element.Formula1?.Text;
        string? formula2 = element.Formula2?.Text;
        (double? minimum, double? maximum) = ReadBounds(type, element.Operator?.InnerText, formula1, formula2);
        return new ExcelDataValidation
        {
            Range = range,
            Type = type,
            Formula1 = formula1,
            Formula2 = formula2,
            Minimum = minimum,
            Maximum = maximum,
            AllowedValues = ReadAllowedValues(type, formula1),
            AllowBlank = element.AllowBlank?.Value ?? false,
            ErrorTitle = element.ErrorTitle?.Value,
            Error = element.Error?.Value,
            PromptTitle = element.PromptTitle?.Value,
            Prompt = element.Prompt?.Value,
        };
    }

    private static (double? Minimum, double? Maximum) ReadBounds(
        ExcelDataValidationType type,
        string? operation,
        string? formula1,
        string? formula2)
    {
        if (type is not (ExcelDataValidationType.WholeNumber or ExcelDataValidationType.DecimalNumber))
        {
            return (null, null);
        }

        if (operation is "between" or "notBetween")
        {
            return (Parse(formula1), Parse(formula2));
        }

        if (operation is "equal" or "notEqual")
        {
            double? value = Parse(formula1);
            return (value, value);
        }

        if (operation is "greaterThan" or "greaterThanOrEqual")
        {
            return (Parse(formula1), null);
        }

        if (operation is "lessThan" or "lessThanOrEqual")
        {
            return (null, Parse(formula1));
        }

        return (null, null);
    }

    private static double? Parse(string? formula)
        => double.TryParse(formula, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value)
            ? value
            : null;

    private static string Format(double value) => value.ToString("R", System.Globalization.CultureInfo.InvariantCulture);

    internal static string FirstArea(string? references)
    {
        if (string.IsNullOrWhiteSpace(references))
        {
            return string.Empty;
        }

        string[] areas = references.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        return areas[0];
    }

    private static string[] ReadAllowedValues(ExcelDataValidationType type, string? formula)
    {
        if (type != ExcelDataValidationType.List || string.IsNullOrEmpty(formula))
        {
            return [];
        }

        string text = formula;
        if (text.Length >= 2 && text[0] == '"' && text[^1] == '"')
        {
            text = text[1..^1];
        }

        if (text.Length == 0)
        {
            return [];
        }

        return text.Split(',');
    }

    private static S.DataValidationValues ToValues(ExcelDataValidationType type) => type switch
    {
        ExcelDataValidationType.None => S.DataValidationValues.None,
        ExcelDataValidationType.WholeNumber => S.DataValidationValues.Whole,
        ExcelDataValidationType.DecimalNumber => S.DataValidationValues.Decimal,
        ExcelDataValidationType.Date => S.DataValidationValues.Date,
        ExcelDataValidationType.Time => S.DataValidationValues.Time,
        ExcelDataValidationType.TextLength => S.DataValidationValues.TextLength,
        ExcelDataValidationType.List => S.DataValidationValues.List,
        ExcelDataValidationType.Custom => S.DataValidationValues.Custom,
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, "The data validation type is not defined."),
    };

    private static ExcelDataValidationType ToType(S.DataValidationValues? value)
    {
        if (value is null || value == S.DataValidationValues.None)
        {
            return ExcelDataValidationType.None;
        }

        if (value == S.DataValidationValues.Whole)
        {
            return ExcelDataValidationType.WholeNumber;
        }

        if (value == S.DataValidationValues.Decimal)
        {
            return ExcelDataValidationType.DecimalNumber;
        }

        if (value == S.DataValidationValues.Date)
        {
            return ExcelDataValidationType.Date;
        }

        if (value == S.DataValidationValues.Time)
        {
            return ExcelDataValidationType.Time;
        }

        if (value == S.DataValidationValues.TextLength)
        {
            return ExcelDataValidationType.TextLength;
        }

        if (value == S.DataValidationValues.List)
        {
            return ExcelDataValidationType.List;
        }

        if (value == S.DataValidationValues.Custom)
        {
            return ExcelDataValidationType.Custom;
        }

        throw new InvalidOfficeFileException($"The data validation type '{value}' is not supported.");
    }
}
