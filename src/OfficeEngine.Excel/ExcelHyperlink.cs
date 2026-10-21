namespace OfficeEngine.Excel;

/// <summary>A hyperlink of a worksheet: the cells it covers and where it points.</summary>
/// <remarks>
/// A hyperlink is either <em>external</em>, with an address that the reader opens in a browser, or
/// <em>internal</em>, with a location inside the workbook. The two are mutually exclusive.
/// </remarks>
public sealed class ExcelHyperlink
{
    internal ExcelHyperlink(
        string reference,
        string? target,
        string? location,
        string? display,
        string? tooltip)
    {
        Reference = reference;
        Target = target;
        Location = location;
        Display = display;
        Tooltip = tooltip;
    }

    /// <summary>Gets the A1-style range the hyperlink covers, for example <c>A1</c>.</summary>
    public string Reference { get; }

    /// <summary>
    /// Gets the address of an external hyperlink, or <see langword="null"/> when the hyperlink points
    /// inside the workbook.
    /// </summary>
    public string? Target { get; }

    /// <summary>
    /// Gets the location of an internal hyperlink, or <see langword="null"/> when the hyperlink is
    /// external.
    /// </summary>
    public string? Location { get; }

    /// <summary>Gets the text that is shown for the hyperlink, when it is set.</summary>
    public string? Display { get; }

    /// <summary>Gets the tooltip of the hyperlink, when it is set.</summary>
    public string? Tooltip { get; }

    /// <summary>Gets whether the hyperlink points outside the workbook.</summary>
    public bool IsExternal => Target is not null;

    /// <inheritdoc />
    public override string ToString() => $"{Reference} → {Target ?? Location}";
}
