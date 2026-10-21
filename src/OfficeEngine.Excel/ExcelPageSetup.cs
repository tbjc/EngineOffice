namespace OfficeEngine.Excel;

/// <summary>The setup of the printed page: orientation, paper size, scale and fit to page.</summary>
public sealed class ExcelPageSetup
{
    internal ExcelPageSetup()
    {
    }

    /// <summary>Initializes a new instance of the <see cref="ExcelPageSetup"/> class.</summary>
    /// <param name="orientation">The orientation of the page.</param>
    /// <param name="paperSize">The paper size.</param>
    /// <param name="scale">
    /// The print scale, as a percentage between 10 and 400. It cannot be combined with
    /// <paramref name="fitToWidth"/> or <paramref name="fitToHeight"/>.
    /// </param>
    /// <param name="fitToWidth">The number of pages the sheet is fitted to horizontally.</param>
    /// <param name="fitToHeight">The number of pages the sheet is fitted to vertically.</param>
    /// <exception cref="ArgumentOutOfRangeException">The scale is out of range or a fit value is negative.</exception>
    /// <exception cref="ArgumentException">A scale and a fit to page value are given at the same time.</exception>
    public ExcelPageSetup(
        ExcelPageOrientation orientation = ExcelPageOrientation.Portrait,
        ExcelPaperSize paperSize = ExcelPaperSize.A4,
        int? scale = null,
        int? fitToWidth = null,
        int? fitToHeight = null)
    {
        if (scale is int percentage && (percentage < 10 || percentage > 400))
        {
            throw new ArgumentOutOfRangeException(
                nameof(scale), percentage, "The print scale must be between 10 and 400.");
        }

        if (fitToWidth is int width && width < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(fitToWidth), width, "The fit to width value cannot be negative.");
        }

        if (fitToHeight is int height && height < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(fitToHeight), height, "The fit to height value cannot be negative.");
        }

        if (scale is not null && (fitToWidth is not null || fitToHeight is not null))
        {
            throw new ArgumentException(
                "The print scale and the fit to page values cannot be used at the same time.",
                nameof(scale));
        }

        Orientation = orientation;
        PaperSize = paperSize;
        Scale = scale;
        FitToWidth = fitToWidth;
        FitToHeight = fitToHeight;
    }

    /// <summary>Gets the orientation of the page.</summary>
    public ExcelPageOrientation Orientation { get; internal init; }

    /// <summary>Gets the paper size.</summary>
    public ExcelPaperSize PaperSize { get; internal init; }

    /// <summary>Gets the print scale as a percentage, or <see langword="null"/> when it is not set.</summary>
    public int? Scale { get; internal init; }

    /// <summary>Gets the number of pages the sheet is fitted to horizontally.</summary>
    public int? FitToWidth { get; internal init; }

    /// <summary>Gets the number of pages the sheet is fitted to vertically.</summary>
    public int? FitToHeight { get; internal init; }

    /// <inheritdoc />
    public override string ToString() => $"{Orientation} {PaperSize}";
}
