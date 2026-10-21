namespace OfficeEngine.Excel;

/// <summary>The four edges of a cell border. A missing edge has no border.</summary>
public readonly struct CellBorders : IEquatable<CellBorders>
{
    /// <summary>Initializes a new instance of the <see cref="CellBorders"/> struct.</summary>
    /// <param name="left">The left edge, or <see langword="null"/> for no left border.</param>
    /// <param name="right">The right edge, or <see langword="null"/> for no right border.</param>
    /// <param name="top">The top edge, or <see langword="null"/> for no top border.</param>
    /// <param name="bottom">The bottom edge, or <see langword="null"/> for no bottom border.</param>
    public CellBorders(
        CellBorder? left = null,
        CellBorder? right = null,
        CellBorder? top = null,
        CellBorder? bottom = null)
    {
        Left = left;
        Right = right;
        Top = top;
        Bottom = bottom;
    }

    /// <summary>Gets the left edge, or <see langword="null"/> when it has no border.</summary>
    public CellBorder? Left { get; }

    /// <summary>Gets the right edge, or <see langword="null"/> when it has no border.</summary>
    public CellBorder? Right { get; }

    /// <summary>Gets the top edge, or <see langword="null"/> when it has no border.</summary>
    public CellBorder? Top { get; }

    /// <summary>Gets the bottom edge, or <see langword="null"/> when it has no border.</summary>
    public CellBorder? Bottom { get; }

    /// <summary>Gets whether all four edges have no border.</summary>
    public bool IsEmpty => Left is null && Right is null && Top is null && Bottom is null;

    /// <inheritdoc />
    public bool Equals(CellBorders other)
        => Left == other.Left && Right == other.Right && Top == other.Top && Bottom == other.Bottom;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is CellBorders other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(Left, Right, Top, Bottom);

    /// <summary>Determines whether two border sets are equal.</summary>
    /// <param name="left">The first border set.</param>
    /// <param name="right">The second border set.</param>
    /// <returns><see langword="true"/> if both border sets are equal; otherwise, <see langword="false"/>.</returns>
    public static bool operator ==(CellBorders left, CellBorders right) => left.Equals(right);

    /// <summary>Determines whether two border sets are different.</summary>
    /// <param name="left">The first border set.</param>
    /// <param name="right">The second border set.</param>
    /// <returns><see langword="true"/> if both border sets are different; otherwise, <see langword="false"/>.</returns>
    public static bool operator !=(CellBorders left, CellBorders right) => !left.Equals(right);
}
