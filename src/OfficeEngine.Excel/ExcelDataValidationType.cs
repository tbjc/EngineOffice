namespace OfficeEngine.Excel;

/// <summary>The kind of rule that a data validation applies.</summary>
public enum ExcelDataValidationType
{
    /// <summary>Any value is accepted.</summary>
    None = 0,

    /// <summary>The value must be a whole number.</summary>
    WholeNumber = 1,

    /// <summary>The value must be a number with decimals.</summary>
    DecimalNumber = 2,

    /// <summary>The value must be a date.</summary>
    Date = 3,

    /// <summary>The value must be a time.</summary>
    Time = 4,

    /// <summary>The value must have a given length.</summary>
    TextLength = 5,

    /// <summary>The value must be one of a list of values.</summary>
    List = 6,

    /// <summary>The value must satisfy a formula.</summary>
    Custom = 7,
}
