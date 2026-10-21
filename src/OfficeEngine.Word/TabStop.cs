using System;
using System.Globalization;

namespace OfficeEngine.Word;

/// <summary>
/// A tab stop of a paragraph: the position where a tab character stops, together with the
/// alignment and the leader used at that position.
/// </summary>
/// <remarks>
/// <para>
/// Office Open XML stores tab stops in <c>w:pPr/w:tabs</c>, one <c>w:tab</c> element per stop, with
/// the position (<c>w:pos</c>) expressed in twentieths of a point. A position is therefore read and
/// written with a resolution of <c>0.05</c> points.
/// </para>
/// <para>
/// Unlike indentation and spacing, the schema (<c>ST_SignedTwipsMeasure</c>) also allows a negative
/// position, so one is accepted here and reproduced faithfully instead of being rejected.
/// </para>
/// <para>
/// Only the position is checked here. An alignment or leader that is not a defined value is
/// rejected by <see cref="Paragraph.AddTabStop(TabStop)"/> before any markup is written, in line
/// with how the rest of the API validates the values it maps.
/// </para>
/// </remarks>
public readonly struct TabStop : IEquatable<TabStop>
{
    /// <summary>Initializes a new instance of the <see cref="TabStop"/> struct.</summary>
    /// <param name="positionPoints">The position of the stop, in points, from the left margin.</param>
    /// <param name="alignment">The alignment used at the stop.</param>
    /// <param name="leader">
    /// The leader character drawn up to the stop. Defaults to <see cref="TabStopLeader.None"/>.
    /// </param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="positionPoints"/> is not a finite number.
    /// </exception>
    public TabStop(double positionPoints, TabStopAlignment alignment, TabStopLeader leader = TabStopLeader.None)
    {
        if (!double.IsFinite(positionPoints))
        {
            throw new ArgumentOutOfRangeException(
                nameof(positionPoints),
                positionPoints,
                "The tab stop position must be a finite number of points.");
        }

        PositionPoints = positionPoints;
        Alignment = alignment;
        Leader = leader;
    }

    /// <summary>Gets the position of the stop, in points, measured from the left margin.</summary>
    public double PositionPoints { get; }

    /// <summary>Gets the alignment used at the stop.</summary>
    public TabStopAlignment Alignment { get; }

    /// <summary>Gets the leader character drawn up to the stop.</summary>
    public TabStopLeader Leader { get; }

    /// <inheritdoc />
    public bool Equals(TabStop other)
        => PositionPoints.Equals(other.PositionPoints) && Alignment == other.Alignment && Leader == other.Leader;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is TabStop other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(PositionPoints, Alignment, Leader);

    /// <summary>Determines whether two tab stops are equal.</summary>
    /// <param name="left">The first tab stop.</param>
    /// <param name="right">The second tab stop.</param>
    /// <returns><see langword="true"/> if both tab stops are equal; otherwise, <see langword="false"/>.</returns>
    public static bool operator ==(TabStop left, TabStop right) => left.Equals(right);

    /// <summary>Determines whether two tab stops are different.</summary>
    /// <param name="left">The first tab stop.</param>
    /// <param name="right">The second tab stop.</param>
    /// <returns><see langword="true"/> if both tab stops are different; otherwise, <see langword="false"/>.</returns>
    public static bool operator !=(TabStop left, TabStop right) => !left.Equals(right);

    /// <inheritdoc />
    public override string ToString()
        => string.Create(CultureInfo.InvariantCulture, $"{Alignment} at {PositionPoints} pt ({Leader})");
}
