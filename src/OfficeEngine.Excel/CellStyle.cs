namespace OfficeEngine.Excel;

/// <summary>
/// A reusable set of cell formatting: font, solid fill, borders and alignment.
/// </summary>
/// <remarks>
/// <para>
/// A style is immutable and can be assigned to any number of cells. Applying the same style to
/// several cells reuses one <c>cellXfs</c> entry when the rest of the cell format is the same.
/// </para>
/// <para>
/// Every property is <em>tri-state</em>: <see langword="null"/> means "use the workbook default"
/// for that property. Assigning a style replaces all the modelled properties of the cell; the number
/// format is separate (<see cref="Cell.NumberFormat"/>) and is preserved. An empty style resets the
/// modelled formatting to the workbook defaults.
/// </para>
/// <para>
/// This version models a solid fill and explicit RGB colours. Pattern fills, theme colours,
/// protection, text rotation, indentation and other formatting this version does not model are not
/// represented here.
/// </para>
/// </remarks>
public sealed record CellStyle
{
    /// <summary>Gets the font name, for example <c>Arial</c>.</summary>
    public string? FontName { get; init; }

    /// <summary>Gets the font size in points.</summary>
    public double? FontSizePoints { get; init; }

    /// <summary>Gets whether the font is bold. <see langword="false"/> explicitly turns it off.</summary>
    public bool? Bold { get; init; }

    /// <summary>Gets whether the font is italic. <see langword="false"/> explicitly turns it off.</summary>
    public bool? Italic { get; init; }

    /// <summary>Gets the underline style.</summary>
    public CellUnderlineStyle? Underline { get; init; }

    /// <summary>Gets the explicit RGB font colour, or <see langword="null"/> for the default colour.</summary>
    public RgbColor? FontColor { get; init; }

    /// <summary>Gets the solid fill colour, or <see langword="null"/> for no fill.</summary>
    public RgbColor? FillColor { get; init; }

    /// <summary>Gets the cell borders, or <see langword="null"/> for no borders.</summary>
    public CellBorders? Borders { get; init; }

    /// <summary>Gets the horizontal alignment.</summary>
    public CellHorizontalAlignment? HorizontalAlignment { get; init; }

    /// <summary>Gets the vertical alignment.</summary>
    public CellVerticalAlignment? VerticalAlignment { get; init; }

    /// <summary>Gets whether text wraps inside the cell.</summary>
    public bool? WrapText { get; init; }
}
