using OfficeEngine.OpenXml;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>The number format resolved for a cell, including whether it represents a date or time.</summary>
internal readonly record struct ResolvedNumberFormat(string? Code, bool IsDateTime);

/// <summary>
/// Reads and writes the subset of the workbook stylesheet needed for number formats and reusable
/// cell styles. Existing style definitions are not rewritten: new components and cell formats are
/// appended and the counts are updated.
/// </summary>
internal static class ExcelStyles
{
    internal const string General = "General";

    private const uint FirstCustomNumberFormatId = 164U;
    private const int MaxNumberFormatLength = 255;
    private const int MaxFontNameLength = 255;
    private const double MaxFontSizePoints = 409.0;

    // Schema order of the children this version models, with the unmodelled children in their
    // expected positions, so a rewritten component keeps its valid element order.
    private static readonly string[] FontElementOrder =
    [
        "b", "i", "strike", "outline", "shadow", "condense", "extend",
        "u", "vertAlign", "sz", "color", "name", "family", "charset", "scheme",
    ];

    private static readonly string[] BorderElementOrder =
    [
        "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal", "start", "end",
    ];

    private static readonly Dictionary<uint, string> BuiltInFormats =
        new()
        {
            [0] = General,
            [1] = "0",
            [2] = "0.00",
            [3] = "#,##0",
            [4] = "#,##0.00",
            [5] = "$#,##0_);($#,##0)",
            [6] = "$#,##0_);[Red]($#,##0)",
            [7] = "$#,##0.00_);($#,##0.00)",
            [8] = "$#,##0.00_);[Red]($#,##0.00)",
            [9] = "0%",
            [10] = "0.00%",
            [11] = "0.00E+00",
            [12] = "# ?/?",
            [13] = "# ??/??",
            [14] = "mm-dd-yy",
            [15] = "d-mmm-yy",
            [16] = "d-mmm",
            [17] = "mmm-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "m/d/yy h:mm",
            [37] = "#,##0 ;(#,##0)",
            [38] = "#,##0 ;[Red](#,##0)",
            [39] = "#,##0.00;(#,##0.00)",
            [40] = "#,##0.00;[Red](#,##0.00)",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [48] = "##0.0E+0",
            [49] = "@",
        };

    private static readonly HashSet<uint> DateTimeBuiltInIds =
    [
        14, 15, 16, 17, 18, 19, 20, 21, 22,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
        45, 46, 47,
        50, 51, 52, 53, 54, 55, 56, 57, 58,
    ];

    private static readonly HashSet<string> ElementsAfterCellFormats =
    [
        "cellStyles",
        "dxfs",
        "tableStyles",
        "colors",
        "extLst",
    ];

    /// <summary>Resolves the effective number format of a cell without changing the workbook.</summary>
    internal static ResolvedNumberFormat Resolve(ExcelWorkbook workbook, S.Cell cell)
    {
        uint styleIndex = cell.StyleIndex?.Value ?? 0U;
        WorkbookStylesPart? part = workbook.WorkbookPart.WorkbookStylesPart;
        S.Stylesheet? stylesheet = part is null
            ? null
            : OoxmlPartReader.Read("the styles part", () => part.Stylesheet);
        if (stylesheet is null)
        {
            if (cell.StyleIndex is not null)
            {
                throw new InvalidOfficeFileException(
                    $"Cell {cell.CellReference?.Value} uses style {styleIndex}, but the workbook has no stylesheet.");
            }

            return new ResolvedNumberFormat(General, false);
        }

        S.CellFormats? cellFormats = stylesheet.CellFormats;
        if (cellFormats is null)
        {
            if (cell.StyleIndex is not null)
            {
                throw new InvalidOfficeFileException(
                    $"Cell {cell.CellReference?.Value} uses style {styleIndex}, but the stylesheet has no cell formats.");
            }

            return new ResolvedNumberFormat(General, false);
        }

        S.CellFormat format = GetCellFormat(cellFormats, styleIndex, cell);
        uint numberFormatId = format.NumberFormatId?.Value ?? 0U;
        if (numberFormatId == 0U)
        {
            return new ResolvedNumberFormat(General, false);
        }

        S.NumberingFormat? custom = stylesheet.NumberingFormats?
            .Elements<S.NumberingFormat>()
            .FirstOrDefault(existing => existing.NumberFormatId?.Value == numberFormatId);
        if (custom is not null)
        {
            string? code = custom.FormatCode?.Value;
            if (string.IsNullOrEmpty(code))
            {
                throw new InvalidOfficeFileException(
                    $"Number format {numberFormatId} does not declare its format code.");
            }

            return new ResolvedNumberFormat(code, NumberFormatCode.IsDateTime(code));
        }

        if (BuiltInFormats.TryGetValue(numberFormatId, out string? builtIn))
        {
            return new ResolvedNumberFormat(builtIn, NumberFormatCode.IsDateTime(builtIn));
        }

        if (DateTimeBuiltInIds.Contains(numberFormatId))
        {
            // Locale-specific built-in date formats have no single canonical code to expose.
            return new ResolvedNumberFormat(null, true);
        }

        if (numberFormatId >= FirstCustomNumberFormatId)
        {
            throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses number format {numberFormatId}, " +
                "which is not defined in the stylesheet.");
        }

        return new ResolvedNumberFormat(null, false);
    }

    /// <summary>Validates a format code and normalizes the General keyword.</summary>
    internal static string Normalize(string? format)
    {
        if (format is null)
        {
            return General;
        }

        if (format.Length == 0)
        {
            throw new ArgumentException("The number format must not be empty.", nameof(format));
        }

        if (format.Length > MaxNumberFormatLength)
        {
            throw new ArgumentException(
                $"The number format must be at most {MaxNumberFormatLength} characters long.",
                nameof(format));
        }

        try
        {
            _ = XmlConvert.VerifyXmlChars(format);
        }
        catch (XmlException ex)
        {
            throw new ArgumentException(
                "The number format contains characters that are not valid in XML.",
                nameof(format),
                ex);
        }

        return string.Equals(format, General, StringComparison.OrdinalIgnoreCase) ? General : format;
    }

    /// <summary>Returns whether a format code contains date or time tokens.</summary>
    internal static bool IsDateTime(string format) => NumberFormatCode.IsDateTime(format);

    /// <summary>Applies an already normalized number format to a cell.</summary>
    internal static void Set(ExcelWorkbook workbook, S.Cell cell, string normalizedFormat)
    {
        if (string.Equals(normalizedFormat, General, StringComparison.Ordinal))
        {
            uint? currentStyleIndex = cell.StyleIndex?.Value;
            S.CellFormats? formats = currentStyleIndex is null
                ? OoxmlPartReader.ReadOrDefault("the styles part", () => workbook.WorkbookPart.WorkbookStylesPart?.Stylesheet)?.CellFormats
                : GetCellFormats(GetStylesheet(workbook));
            if (formats is null)
            {
                return;
            }

            S.CellFormat current = GetCellFormat(formats, currentStyleIndex ?? 0U, cell);
            if ((current.NumberFormatId?.Value ?? 0U) == 0U)
            {
                return;
            }

            S.CellFormat clone = (S.CellFormat)current.CloneNode(true);
            clone.NumberFormatId = 0U;
            clone.ApplyNumberFormat = false;
            cell.StyleIndex = GetOrAddCellFormat(formats, clone);
            return;
        }

        S.Stylesheet targetStylesheet = GetOrCreateStylesheet(workbook, cell);
        S.CellFormats targetFormats = cell.StyleIndex is null
            ? GetOrCreateCellFormats(targetStylesheet)
            : GetCellFormats(targetStylesheet);
        uint numberFormatId = GetOrCreateNumberFormatId(targetStylesheet, normalizedFormat);
        uint currentIndex = cell.StyleIndex?.Value ?? 0U;
        S.CellFormat currentFormat = GetCellFormat(targetFormats, currentIndex, cell);
        if ((currentFormat.NumberFormatId?.Value ?? 0U) == numberFormatId &&
            currentFormat.ApplyNumberFormat?.Value != false)
        {
            return;
        }

        S.CellFormat cloneFormat = (S.CellFormat)currentFormat.CloneNode(true);
        cloneFormat.NumberFormatId = numberFormatId;
        cloneFormat.ApplyNumberFormat = true;
        cell.StyleIndex = GetOrAddCellFormat(targetFormats, cloneFormat);
    }

    /// <summary>Resolves the modelled cell style without changing the workbook.</summary>
    internal static CellStyle ResolveStyle(ExcelWorkbook workbook, S.Cell cell)
    {
        uint styleIndex = cell.StyleIndex?.Value ?? 0U;
        WorkbookStylesPart? part = workbook.WorkbookPart.WorkbookStylesPart;
        S.Stylesheet? stylesheet = part is null
            ? null
            : OoxmlPartReader.Read("the styles part", () => part.Stylesheet);
        if (stylesheet is null)
        {
            if (cell.StyleIndex is not null)
            {
                throw new InvalidOfficeFileException(
                    $"Cell {cell.CellReference?.Value} uses style {styleIndex}, but the workbook has no stylesheet.");
            }

            return new CellStyle();
        }

        S.CellFormats? cellFormats = stylesheet.CellFormats;
        if (cellFormats is null)
        {
            if (cell.StyleIndex is not null)
            {
                throw new InvalidOfficeFileException(
                    $"Cell {cell.CellReference?.Value} uses style {styleIndex}, but the stylesheet has no cell formats.");
            }

            return new CellStyle();
        }

        S.CellFormat format = GetCellFormat(cellFormats, styleIndex, cell);
        S.Font font = GetFont(stylesheet, format.FontId?.Value ?? 0U, cell);
        S.Fill fill = GetFill(stylesheet, format.FillId?.Value ?? 0U, cell);
        S.Border border = GetBorder(stylesheet, format.BorderId?.Value ?? 0U, cell);

        return new CellStyle
        {
            FontName = ReadFontName(font),
            FontSizePoints = ReadFontSize(font),
            Bold = ReadBoolean(font.Bold),
            Italic = ReadBoolean(font.Italic),
            Underline = ReadUnderline(font.Underline),
            FontColor = ReadColor(font.Color),
            FillColor = ReadFillColor(fill),
            Borders = ReadBorders(border),
            HorizontalAlignment = ReadHorizontalAlignment(format.Alignment),
            VerticalAlignment = ReadVerticalAlignment(format.Alignment),
            WrapText = format.Alignment?.WrapText?.Value,
        };
    }

    /// <summary>Validates a style before it touches the workbook.</summary>
    internal static void ValidateStyle(CellStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);

        if (style.FontName is string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName))
            {
                throw new ArgumentException("The font name must not be empty.", nameof(style));
            }

            if (fontName.Length > MaxFontNameLength)
            {
                throw new ArgumentException(
                    $"The font name must be at most {MaxFontNameLength} characters long.",
                    nameof(style));
            }

            try
            {
                _ = XmlConvert.VerifyXmlChars(fontName);
            }
            catch (XmlException ex)
            {
                throw new ArgumentException(
                    "The font name contains characters that are not valid in XML.",
                    nameof(style),
                    ex);
            }
        }

        if (style.FontSizePoints is double fontSize &&
            (!double.IsFinite(fontSize) || fontSize <= 0.0 || fontSize > MaxFontSizePoints))
        {
            throw new ArgumentOutOfRangeException(
                nameof(style),
                fontSize,
                $"The font size must be greater than 0 and at most {MaxFontSizePoints} points.");
        }

        if (style.Underline is CellUnderlineStyle underline && !Enum.IsDefined(underline))
        {
            throw new ArgumentOutOfRangeException(nameof(style), underline, "The underline style is not a defined value.");
        }

        if (style.HorizontalAlignment is CellHorizontalAlignment horizontal && !Enum.IsDefined(horizontal))
        {
            throw new ArgumentOutOfRangeException(
                nameof(style),
                horizontal,
                "The horizontal alignment is not a defined value.");
        }

        if (style.VerticalAlignment is CellVerticalAlignment vertical && !Enum.IsDefined(vertical))
        {
            throw new ArgumentOutOfRangeException(
                nameof(style),
                vertical,
                "The vertical alignment is not a defined value.");
        }
    }

    /// <summary>Gets whether a style resets every modelled property to its default.</summary>
    internal static bool IsDefaultStyle(CellStyle style)
        => style.FontName is null
            && style.FontSizePoints is null
            && style.Bold is null
            && style.Italic is null
            && style.Underline is null
            && style.FontColor is null
            && style.FillColor is null
            && (style.Borders is not CellBorders borders || borders.IsEmpty)
            && style.HorizontalAlignment is null
            && style.VerticalAlignment is null
            && style.WrapText is null;

    /// <summary>Applies a validated style to a cell, preserving its number format and unmodelled xf data.</summary>
    internal static void ApplyStyle(ExcelWorkbook workbook, S.Cell cell, CellStyle style)
    {
        bool isDefault = IsDefaultStyle(style);
        if (isDefault && cell.StyleIndex is null)
        {
            S.CellFormats? existingFormats = OoxmlPartReader.ReadOrDefault(
            "the styles part",
            () => workbook.WorkbookPart.WorkbookStylesPart?.Stylesheet)?.CellFormats;
            if (existingFormats is null)
            {
                return;
            }

            if (IsDefaultFormat(GetCellFormat(existingFormats, 0U, cell)))
            {
                return;
            }
        }

        S.Stylesheet stylesheet = GetOrCreateStylesheet(workbook, cell);
        S.CellFormats formats = cell.StyleIndex is null
            ? GetOrCreateCellFormats(stylesheet)
            : GetCellFormats(stylesheet);
        uint currentIndex = cell.StyleIndex?.Value ?? 0U;
        S.CellFormat current = GetCellFormat(formats, currentIndex, cell);

        uint fontId = 0U;
        uint fillId = 0U;
        uint borderId = 0U;
        S.Alignment? alignment = null;
        if (!isDefault)
        {
            S.Font currentFont = GetFont(stylesheet, current.FontId?.Value ?? 0U, cell);
            S.Border currentBorder = GetBorder(stylesheet, current.BorderId?.Value ?? 0U, cell);
            fontId = GetOrAddFont(stylesheet, style, currentFont);
            fillId = GetOrAddFill(stylesheet, style.FillColor);
            borderId = GetOrAddBorder(stylesheet, style.Borders, currentBorder);
            alignment = CreateAlignment(style, current.Alignment);
        }

        S.CellFormat clone = (S.CellFormat)current.CloneNode(true);
        clone.FontId = fontId;
        clone.FillId = fillId;
        clone.BorderId = borderId;
        clone.ApplyFont = true;
        clone.ApplyFill = true;
        clone.ApplyBorder = true;
        clone.ApplyAlignment = true;
        clone.RemoveAllChildren<S.Alignment>();
        if (alignment is not null)
        {
            InsertAlignment(clone, alignment);
        }

        cell.StyleIndex = GetOrAddCellFormat(formats, clone);
    }

    private static bool IsDefaultFormat(S.CellFormat format)
        => (format.FontId?.Value ?? 0U) == 0U
            && (format.FillId?.Value ?? 0U) == 0U
            && (format.BorderId?.Value ?? 0U) == 0U
            && format.Alignment is null;

    private static S.Font GetFont(S.Stylesheet stylesheet, uint index, S.Cell cell)
    {
        S.Fonts? fonts = stylesheet.Fonts;
        if (fonts is null)
        {
            if (index == 0U)
            {
                return new S.Font();
            }

            throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses font {index}, but the stylesheet has no fonts.");
        }

        if (index > int.MaxValue)
        {
            throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses font {index}, which is outside the supported range.");
        }

        return fonts.Elements<S.Font>().ElementAtOrDefault((int)index)
            ?? throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses font {index}, which does not exist.");
    }

    private static S.Fill GetFill(S.Stylesheet stylesheet, uint index, S.Cell cell)
    {
        S.Fills? fills = stylesheet.Fills;
        if (fills is null)
        {
            if (index == 0U)
            {
                return new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None });
            }

            throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses fill {index}, but the stylesheet has no fills.");
        }

        if (index > int.MaxValue)
        {
            throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses fill {index}, which is outside the supported range.");
        }

        return fills.Elements<S.Fill>().ElementAtOrDefault((int)index)
            ?? throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses fill {index}, which does not exist.");
    }

    private static S.Border GetBorder(S.Stylesheet stylesheet, uint index, S.Cell cell)
    {
        S.Borders? borders = stylesheet.Borders;
        if (borders is null)
        {
            if (index == 0U)
            {
                return new S.Border();
            }

            throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses border {index}, but the stylesheet has no borders.");
        }

        if (index > int.MaxValue)
        {
            throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses border {index}, which is outside the supported range.");
        }

        return borders.Elements<S.Border>().ElementAtOrDefault((int)index)
            ?? throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses border {index}, which does not exist.");
    }

    private static string? ReadFontName(S.Font font)
    {
        string? value = font.FontName?.Val?.Value;
        return string.IsNullOrEmpty(value) ? null : value;
    }

    private static double? ReadFontSize(S.Font font)
    {
        double? value = font.FontSize?.Val?.Value;
        if (value is double size && (!double.IsFinite(size) || size <= 0.0))
        {
            throw new InvalidOfficeFileException("A font has an invalid size.");
        }

        return value;
    }

    private static bool? ReadBoolean(S.Bold? bold) => bold is null ? null : bold.Val?.Value ?? true;

    private static bool? ReadBoolean(S.Italic? italic) => italic is null ? null : italic.Val?.Value ?? true;

    private static CellUnderlineStyle? ReadUnderline(S.Underline? underline)
    {
        if (underline is null)
        {
            return null;
        }

        S.UnderlineValues value = underline.Val?.Value ?? S.UnderlineValues.Single;
        if (value == S.UnderlineValues.None)
        {
            return CellUnderlineStyle.None;
        }

        if (value == S.UnderlineValues.Single)
        {
            return CellUnderlineStyle.SingleLine;
        }

        if (value == S.UnderlineValues.Double)
        {
            return CellUnderlineStyle.DoubleLine;
        }

        if (value == S.UnderlineValues.SingleAccounting)
        {
            return CellUnderlineStyle.SingleAccounting;
        }

        if (value == S.UnderlineValues.DoubleAccounting)
        {
            return CellUnderlineStyle.DoubleAccounting;
        }

        throw new InvalidOfficeFileException($"The underline style '{value}' is not supported.");
    }

    private static RgbColor? ReadColor(S.Color? color) => ReadColor(color?.Rgb?.Value);

    private static RgbColor? ReadColor(S.ForegroundColor? color) => ReadColor(color?.Rgb?.Value);

    private static RgbColor? ReadColor(string? rgb)
    {
        if (string.IsNullOrEmpty(rgb))
        {
            return null;
        }

        // SpreadsheetML stores colours as ARGB; this version models the RGB part only.
        string value = rgb.Length == 8 ? rgb[2..] : rgb;
        if (!RgbColor.TryParseHex(value, out RgbColor parsed))
        {
            throw new InvalidOfficeFileException($"The colour '{rgb}' is not an RGB or ARGB value.");
        }

        return parsed;
    }

    private static string ToExcelColor(RgbColor color) => string.Concat("FF", color.ToHex());

    private static RgbColor? ReadFillColor(S.Fill fill)
    {
        S.PatternFill? pattern = fill.PatternFill;
        if (pattern is null)
        {
            return null;
        }

        S.PatternValues? type = pattern.PatternType?.Value;
        if (type is not null && type != S.PatternValues.Solid)
        {
            return null;
        }

        return ReadColor(pattern.ForegroundColor);
    }

    private static CellBorders? ReadBorders(S.Border border)
    {
        CellBorder? left = ReadBorderEdge(border.LeftBorder);
        CellBorder? right = ReadBorderEdge(border.RightBorder);
        CellBorder? top = ReadBorderEdge(border.TopBorder);
        CellBorder? bottom = ReadBorderEdge(border.BottomBorder);

        return left is null && right is null && top is null && bottom is null
            ? null
            : new CellBorders(left, right, top, bottom);
    }

    private static CellBorder? ReadBorderEdge(S.BorderPropertiesType? edge)
    {
        if (edge is null)
        {
            return null;
        }

        S.BorderStyleValues? style = edge.Style?.Value;
        if (style is null || style == S.BorderStyleValues.None)
        {
            return null;
        }

        return new CellBorder(ToCellBorderStyle(style.Value), ReadColor(edge.Color));
    }

    private static CellHorizontalAlignment? ReadHorizontalAlignment(S.Alignment? alignment)
    {
        if (alignment?.Horizontal?.Value is not S.HorizontalAlignmentValues value)
        {
            return null;
        }

        if (value == S.HorizontalAlignmentValues.General)
        {
            return CellHorizontalAlignment.General;
        }

        if (value == S.HorizontalAlignmentValues.Left)
        {
            return CellHorizontalAlignment.Left;
        }

        if (value == S.HorizontalAlignmentValues.Center)
        {
            return CellHorizontalAlignment.Center;
        }

        if (value == S.HorizontalAlignmentValues.Right)
        {
            return CellHorizontalAlignment.Right;
        }

        if (value == S.HorizontalAlignmentValues.Fill)
        {
            return CellHorizontalAlignment.Fill;
        }

        if (value == S.HorizontalAlignmentValues.Justify)
        {
            return CellHorizontalAlignment.Justify;
        }

        if (value == S.HorizontalAlignmentValues.CenterContinuous)
        {
            return CellHorizontalAlignment.CenterContinuous;
        }

        if (value == S.HorizontalAlignmentValues.Distributed)
        {
            return CellHorizontalAlignment.Distributed;
        }

        throw new InvalidOfficeFileException($"The horizontal alignment '{value}' is not supported.");
    }

    private static CellVerticalAlignment? ReadVerticalAlignment(S.Alignment? alignment)
    {
        if (alignment?.Vertical?.Value is not S.VerticalAlignmentValues value)
        {
            return null;
        }

        if (value == S.VerticalAlignmentValues.Top)
        {
            return CellVerticalAlignment.Top;
        }

        if (value == S.VerticalAlignmentValues.Center)
        {
            return CellVerticalAlignment.Center;
        }

        if (value == S.VerticalAlignmentValues.Bottom)
        {
            return CellVerticalAlignment.Bottom;
        }

        if (value == S.VerticalAlignmentValues.Justify)
        {
            return CellVerticalAlignment.Justify;
        }

        if (value == S.VerticalAlignmentValues.Distributed)
        {
            return CellVerticalAlignment.Distributed;
        }

        throw new InvalidOfficeFileException($"The vertical alignment '{value}' is not supported.");
    }

    private static S.Font CreateFont(CellStyle style, S.Font current)
    {
        S.Font font = (S.Font)current.CloneNode(true);
        font.RemoveAllChildren<S.Bold>();
        font.RemoveAllChildren<S.Italic>();
        font.RemoveAllChildren<S.Underline>();
        font.RemoveAllChildren<S.FontSize>();
        font.RemoveAllChildren<S.Color>();
        font.RemoveAllChildren<S.FontName>();

        if (style.Bold is bool bold)
        {
            InsertInOrder(font, new S.Bold { Val = bold }, FontElementOrder);
        }

        if (style.Italic is bool italic)
        {
            InsertInOrder(font, new S.Italic { Val = italic }, FontElementOrder);
        }

        if (style.Underline is CellUnderlineStyle underline)
        {
            InsertInOrder(font, new S.Underline { Val = ToUnderline(underline) }, FontElementOrder);
        }

        if (style.FontSizePoints is double fontSize)
        {
            InsertInOrder(font, new S.FontSize { Val = fontSize }, FontElementOrder);
        }

        if (style.FontColor is RgbColor fontColor)
        {
            InsertInOrder(font, new S.Color { Rgb = ToExcelColor(fontColor) }, FontElementOrder);
        }

        if (style.FontName is string fontName)
        {
            InsertInOrder(font, new S.FontName { Val = fontName }, FontElementOrder);
        }

        return font;
    }

    private static S.Fill CreateFill(RgbColor? color)
    {
        if (color is not RgbColor fillColor)
        {
            return new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None });
        }

        return new S.Fill(new S.PatternFill
        {
            PatternType = S.PatternValues.Solid,
            ForegroundColor = new S.ForegroundColor { Rgb = ToExcelColor(fillColor) },
        });
    }

    private static S.Border CreateBorder(CellBorders? borders, S.Border current)
    {
        S.Border border = (S.Border)current.CloneNode(true);
        border.RemoveAllChildren<S.LeftBorder>();
        border.RemoveAllChildren<S.RightBorder>();
        border.RemoveAllChildren<S.TopBorder>();
        border.RemoveAllChildren<S.BottomBorder>();

        if (borders?.Left is CellBorder left)
        {
            InsertInOrder(border, CreateBorderEdge<S.LeftBorder>(left), BorderElementOrder);
        }

        if (borders?.Right is CellBorder right)
        {
            InsertInOrder(border, CreateBorderEdge<S.RightBorder>(right), BorderElementOrder);
        }

        if (borders?.Top is CellBorder top)
        {
            InsertInOrder(border, CreateBorderEdge<S.TopBorder>(top), BorderElementOrder);
        }

        if (borders?.Bottom is CellBorder bottom)
        {
            InsertInOrder(border, CreateBorderEdge<S.BottomBorder>(bottom), BorderElementOrder);
        }

        return border;
    }

    private static T CreateBorderEdge<T>(CellBorder border)
        where T : S.BorderPropertiesType, new()
    {
        T edge = new() { Style = ToBorderStyle(border.Style) };
        if (border.Color is RgbColor color)
        {
            edge.Color = new S.Color { Rgb = ToExcelColor(color) };
        }

        return edge;
    }

    private static S.Alignment? CreateAlignment(CellStyle style, S.Alignment? current)
    {
        S.Alignment alignment = current is null
            ? new S.Alignment()
            : (S.Alignment)current.CloneNode(true);
        if (style.HorizontalAlignment is CellHorizontalAlignment horizontal)
        {
            alignment.Horizontal = ToHorizontalAlignment(horizontal);
        }
        else
        {
            alignment.Horizontal = null;
        }

        if (style.VerticalAlignment is CellVerticalAlignment vertical)
        {
            alignment.Vertical = ToVerticalAlignment(vertical);
        }
        else
        {
            alignment.Vertical = null;
        }

        if (style.WrapText is bool wrapText)
        {
            alignment.WrapText = wrapText;
        }
        else
        {
            alignment.WrapText = null;
        }

        return alignment.GetAttributes().Count == 0 && alignment.ChildElements.Count == 0
            ? null
            : alignment;
    }

    private static void InsertAlignment(S.CellFormat format, S.Alignment alignment)
    {
        OpenXmlElement? next = format.ChildElements
            .FirstOrDefault(child => child.LocalName is "protection" or "extLst");
        if (next is null)
        {
            format.AppendChild(alignment);
        }
        else
        {
            format.InsertBefore(alignment, next);
        }
    }

    private static void InsertInOrder(OpenXmlElement parent, OpenXmlElement child, string[] order)
    {
        int childOrder = Array.IndexOf(order, child.LocalName);
        OpenXmlElement? next = parent.ChildElements
            .FirstOrDefault(existing => Array.IndexOf(order, existing.LocalName) > childOrder);
        if (next is null)
        {
            parent.AppendChild(child);
        }
        else
        {
            parent.InsertBefore(child, next);
        }
    }

    private static uint GetOrAddFont(S.Stylesheet stylesheet, CellStyle style, S.Font current)
    {
        S.Fonts fonts = stylesheet.Fonts
            ?? throw new InvalidOfficeFileException("The stylesheet does not contain fonts.");
        S.Font candidate = CreateFont(style, current);
        string candidateXml = candidate.OuterXml;
        uint index = 0U;
        foreach (S.Font existing in fonts.Elements<S.Font>())
        {
            if (string.Equals(existing.OuterXml, candidateXml, StringComparison.Ordinal))
            {
                return index;
            }

            index++;
        }

        fonts.AppendChild(candidate);
        fonts.Count = (uint)fonts.Elements<S.Font>().Count();
        return index;
    }

    private static uint GetOrAddFill(S.Stylesheet stylesheet, RgbColor? color)
    {
        S.Fills fills = stylesheet.Fills
            ?? throw new InvalidOfficeFileException("The stylesheet does not contain fills.");
        S.Fill candidate = CreateFill(color);
        string candidateXml = candidate.OuterXml;
        uint index = 0U;
        foreach (S.Fill existing in fills.Elements<S.Fill>())
        {
            if (string.Equals(existing.OuterXml, candidateXml, StringComparison.Ordinal))
            {
                return index;
            }

            index++;
        }

        fills.AppendChild(candidate);
        fills.Count = (uint)fills.Elements<S.Fill>().Count();
        return index;
    }

    private static uint GetOrAddBorder(S.Stylesheet stylesheet, CellBorders? borders, S.Border current)
    {
        S.Borders collection = stylesheet.Borders
            ?? throw new InvalidOfficeFileException("The stylesheet does not contain borders.");
        S.Border candidate = CreateBorder(borders, current);
        string candidateXml = candidate.OuterXml;
        uint index = 0U;
        foreach (S.Border existing in collection.Elements<S.Border>())
        {
            if (string.Equals(existing.OuterXml, candidateXml, StringComparison.Ordinal))
            {
                return index;
            }

            index++;
        }

        collection.AppendChild(candidate);
        collection.Count = (uint)collection.Elements<S.Border>().Count();
        return index;
    }

    private static S.BorderStyleValues ToBorderStyle(CellBorderStyle style) => style switch
    {
        CellBorderStyle.None => S.BorderStyleValues.None,
        CellBorderStyle.Thin => S.BorderStyleValues.Thin,
        CellBorderStyle.Medium => S.BorderStyleValues.Medium,
        CellBorderStyle.Dashed => S.BorderStyleValues.Dashed,
        CellBorderStyle.Dotted => S.BorderStyleValues.Dotted,
        CellBorderStyle.Thick => S.BorderStyleValues.Thick,
        CellBorderStyle.DoubleLine => S.BorderStyleValues.Double,
        CellBorderStyle.Hair => S.BorderStyleValues.Hair,
        CellBorderStyle.MediumDashed => S.BorderStyleValues.MediumDashed,
        CellBorderStyle.DashDot => S.BorderStyleValues.DashDot,
        CellBorderStyle.MediumDashDot => S.BorderStyleValues.MediumDashDot,
        CellBorderStyle.DashDotDot => S.BorderStyleValues.DashDotDot,
        CellBorderStyle.MediumDashDotDot => S.BorderStyleValues.MediumDashDotDot,
        CellBorderStyle.SlantDashDot => S.BorderStyleValues.SlantDashDot,
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, "The cell border style is not defined."),
    };

    private static CellBorderStyle ToCellBorderStyle(S.BorderStyleValues style)
    {
        if (style == S.BorderStyleValues.None)
        {
            return CellBorderStyle.None;
        }

        if (style == S.BorderStyleValues.Thin)
        {
            return CellBorderStyle.Thin;
        }

        if (style == S.BorderStyleValues.Medium)
        {
            return CellBorderStyle.Medium;
        }

        if (style == S.BorderStyleValues.Dashed)
        {
            return CellBorderStyle.Dashed;
        }

        if (style == S.BorderStyleValues.Dotted)
        {
            return CellBorderStyle.Dotted;
        }

        if (style == S.BorderStyleValues.Thick)
        {
            return CellBorderStyle.Thick;
        }

        if (style == S.BorderStyleValues.Double)
        {
            return CellBorderStyle.DoubleLine;
        }

        if (style == S.BorderStyleValues.Hair)
        {
            return CellBorderStyle.Hair;
        }

        if (style == S.BorderStyleValues.MediumDashed)
        {
            return CellBorderStyle.MediumDashed;
        }

        if (style == S.BorderStyleValues.DashDot)
        {
            return CellBorderStyle.DashDot;
        }

        if (style == S.BorderStyleValues.MediumDashDot)
        {
            return CellBorderStyle.MediumDashDot;
        }

        if (style == S.BorderStyleValues.DashDotDot)
        {
            return CellBorderStyle.DashDotDot;
        }

        if (style == S.BorderStyleValues.MediumDashDotDot)
        {
            return CellBorderStyle.MediumDashDotDot;
        }

        if (style == S.BorderStyleValues.SlantDashDot)
        {
            return CellBorderStyle.SlantDashDot;
        }

        throw new InvalidOfficeFileException($"The cell border style '{style}' is not supported.");
    }

    private static S.UnderlineValues ToUnderline(CellUnderlineStyle underline) => underline switch
    {
        CellUnderlineStyle.None => S.UnderlineValues.None,
        CellUnderlineStyle.SingleLine => S.UnderlineValues.Single,
        CellUnderlineStyle.DoubleLine => S.UnderlineValues.Double,
        CellUnderlineStyle.SingleAccounting => S.UnderlineValues.SingleAccounting,
        CellUnderlineStyle.DoubleAccounting => S.UnderlineValues.DoubleAccounting,
        _ => throw new ArgumentOutOfRangeException(nameof(underline), underline, "The underline style is not defined."),
    };

    private static S.HorizontalAlignmentValues ToHorizontalAlignment(CellHorizontalAlignment alignment) => alignment switch
    {
        CellHorizontalAlignment.General => S.HorizontalAlignmentValues.General,
        CellHorizontalAlignment.Left => S.HorizontalAlignmentValues.Left,
        CellHorizontalAlignment.Center => S.HorizontalAlignmentValues.Center,
        CellHorizontalAlignment.Right => S.HorizontalAlignmentValues.Right,
        CellHorizontalAlignment.Fill => S.HorizontalAlignmentValues.Fill,
        CellHorizontalAlignment.Justify => S.HorizontalAlignmentValues.Justify,
        CellHorizontalAlignment.CenterContinuous => S.HorizontalAlignmentValues.CenterContinuous,
        CellHorizontalAlignment.Distributed => S.HorizontalAlignmentValues.Distributed,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, "The horizontal alignment is not defined."),
    };

    private static S.VerticalAlignmentValues ToVerticalAlignment(CellVerticalAlignment alignment) => alignment switch
    {
        CellVerticalAlignment.Top => S.VerticalAlignmentValues.Top,
        CellVerticalAlignment.Center => S.VerticalAlignmentValues.Center,
        CellVerticalAlignment.Bottom => S.VerticalAlignmentValues.Bottom,
        CellVerticalAlignment.Justify => S.VerticalAlignmentValues.Justify,
        CellVerticalAlignment.Distributed => S.VerticalAlignmentValues.Distributed,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, "The vertical alignment is not defined."),
    };

    private static S.Stylesheet GetStylesheet(ExcelWorkbook workbook)
        => workbook.WorkbookPart.WorkbookStylesPart?.Stylesheet
           ?? throw new InvalidOfficeFileException("The workbook has a styled cell but no readable stylesheet.");

    private static S.Stylesheet GetOrCreateStylesheet(ExcelWorkbook workbook, S.Cell cell)
    {
        WorkbookPart workbookPart = workbook.WorkbookPart;
        WorkbookStylesPart? existing = workbookPart.WorkbookStylesPart;
        if (existing?.Stylesheet is not null)
        {
            return existing.Stylesheet;
        }

        if (cell.StyleIndex?.Value is uint styleIndex && styleIndex != 0U)
        {
            throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses style {styleIndex}, but the workbook has no stylesheet.");
        }

        WorkbookStylesPart part = existing ?? workbookPart.AddNewPart<WorkbookStylesPart>();
        part.Stylesheet = CreateStylesheet();
        return part.Stylesheet;
    }

    private static S.Stylesheet CreateStylesheet()
        => new(
            new S.Fonts(new S.Font()) { Count = 1U },
            new S.Fills(new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None })) { Count = 1U },
            new S.Borders(new S.Border()) { Count = 1U },
            new S.CellStyleFormats(
                new S.CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U })
            { Count = 1U },
            new S.CellFormats(
                new S.CellFormat
                {
                    NumberFormatId = 0U,
                    FontId = 0U,
                    FillId = 0U,
                    BorderId = 0U,
                    FormatId = 0U,
                })
            { Count = 1U },
            new S.CellStyles(new S.CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U })
            { Count = 1U });

    private static S.CellFormats GetCellFormats(S.Stylesheet stylesheet)
        => stylesheet.CellFormats
           ?? throw new InvalidOfficeFileException("The stylesheet does not contain cell formats.");

    private static S.CellFormats GetOrCreateCellFormats(S.Stylesheet stylesheet)
    {
        S.CellFormats? formats = stylesheet.CellFormats;
        if (formats is not null)
        {
            return formats;
        }

        if (stylesheet.Fonts is null ||
            stylesheet.Fills is null ||
            stylesheet.Borders is null ||
            stylesheet.CellStyleFormats is null)
        {
            throw new InvalidOfficeFileException(
                "The stylesheet does not contain the default font, fill, border and cell style " +
                "formats required to add a cell format.");
        }

        formats = new S.CellFormats(
            new S.CellFormat
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
            })
        { Count = 1U };

        OpenXmlElement? next = stylesheet.ChildElements
            .FirstOrDefault(child => ElementsAfterCellFormats.Contains(child.LocalName));
        if (next is null)
        {
            stylesheet.AppendChild(formats);
        }
        else
        {
            stylesheet.InsertBefore(formats, next);
        }

        return formats;
    }

    private static S.CellFormat GetCellFormat(S.CellFormats formats, uint index, S.Cell cell)
    {
        if (index > int.MaxValue)
        {
            throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses style {index}, which is outside the supported range.");
        }

        S.CellFormat? format = formats.Elements<S.CellFormat>().ElementAtOrDefault((int)index);
        return format
            ?? throw new InvalidOfficeFileException(
                $"Cell {cell.CellReference?.Value} uses style {index}, which does not exist.");
    }

    private static uint GetOrAddCellFormat(S.CellFormats formats, S.CellFormat candidate)
    {
        string candidateXml = candidate.OuterXml;
        uint index = 0U;
        foreach (S.CellFormat existing in formats.Elements<S.CellFormat>())
        {
            if (string.Equals(existing.OuterXml, candidateXml, StringComparison.Ordinal))
            {
                return index;
            }

            index++;
        }

        formats.AppendChild(candidate);
        formats.Count = (uint)formats.Elements<S.CellFormat>().Count();
        return index;
    }

    private static uint GetOrCreateNumberFormatId(S.Stylesheet stylesheet, string format)
    {
        S.NumberingFormats formats = stylesheet.NumberingFormats
            ?? CreateNumberingFormats(stylesheet);

        foreach (S.NumberingFormat existing in formats.Elements<S.NumberingFormat>())
        {
            if (string.Equals(existing.FormatCode?.Value, format, StringComparison.Ordinal))
            {
                return existing.NumberFormatId?.Value
                    ?? throw new InvalidOfficeFileException("A custom number format does not declare its id.");
            }
        }

        uint next = FirstCustomNumberFormatId;
        foreach (S.NumberingFormat existing in formats.Elements<S.NumberingFormat>())
        {
            if (existing.NumberFormatId?.Value is uint id && id >= next)
            {
                next = id + 1U;
            }
        }

        if (next > ushort.MaxValue)
        {
            throw new UnsupportedOfficeFeatureException(
                "The workbook has no room for another custom number format.");
        }

        formats.AppendChild(new S.NumberingFormat { NumberFormatId = next, FormatCode = format });
        formats.Count = (uint)formats.Elements<S.NumberingFormat>().Count();
        return next;
    }

    private static S.NumberingFormats CreateNumberingFormats(S.Stylesheet stylesheet)
    {
        S.NumberingFormats formats = new() { Count = 0U };
        OpenXmlElement? first = stylesheet.FirstChild;
        if (first is null)
        {
            stylesheet.AppendChild(formats);
        }
        else
        {
            stylesheet.InsertBefore(formats, first);
        }

        return formats;
    }
}

/// <summary>
/// Detects date/time tokens in an Excel number format code, ignoring quoted literals, escaped
/// characters, colors, conditions and locale identifiers in brackets.
/// </summary>
internal static class NumberFormatCode
{
    internal static bool IsDateTime(string format)
    {
        for (int i = 0; i < format.Length; i++)
        {
            if (format.AsSpan(i).StartsWith(ExcelStyles.General, StringComparison.OrdinalIgnoreCase))
            {
                i += ExcelStyles.General.Length - 1;
                continue;
            }

            char current = format[i];
            switch (current)
            {
                case '"':
                    i = SkipQuotedLiteral(format, i);
                    if (i < 0)
                    {
                        return false;
                    }

                    break;

                case '\\':
                case '_':
                case '*':
                    i++;
                    break;

                case '[':
                    int end = format.IndexOf(']', i + 1);
                    if (end < 0)
                    {
                        return false;
                    }

                    if (IsElapsedTime(format.AsSpan(i + 1, end - i - 1)))
                    {
                        return true;
                    }

                    i = end;
                    break;

                case 'y' or 'Y' or 'm' or 'M' or 'd' or 'D' or
                     'h' or 'H' or 's' or 'S' or 'b' or 'B' or 'g' or 'G':
                    return true;

                case 'e' or 'E':
                    if (i + 1 < format.Length &&
                        (format[i + 1] is '+' or '-' or >= '0' and <= '9'))
                    {
                        i++;
                    }
                    else
                    {
                        return true;
                    }

                    break;

                case 'a' or 'A':
                    if (i + 2 < format.Length &&
                        format[i + 1] == '/' &&
                        (format[i + 2] is 'p' or 'P'))
                    {
                        return true;
                    }

                    break;
            }
        }

        return false;
    }

    private static bool IsElapsedTime(ReadOnlySpan<char> content)
    {
        if (content.IsEmpty)
        {
            return false;
        }

        foreach (char character in content)
        {
            if (character is not ('h' or 'H' or 'm' or 'M' or 's' or 'S'))
            {
                return false;
            }
        }

        return true;
    }

    private static int SkipQuotedLiteral(string format, int start)
    {
        for (int i = start + 1; i < format.Length; i++)
        {
            if (format[i] != '"')
            {
                continue;
            }

            if (i + 1 < format.Length && format[i + 1] == '"')
            {
                i++;
                continue;
            }

            return i;
        }

        return -1;
    }
}
