using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>Finds and replaces text inside the runs of a paragraph.</summary>
/// <remarks>
/// The searched text is the concatenation of the <c>w:t</c> elements of the direct runs of the
/// paragraph and of the runs of its direct hyperlinks. Tabs, breaks, drawings and the text of
/// structures this version does not model (content controls, nested tables) do not take part.
/// </remarks>
internal static class WordTextSearch
{
    internal static IReadOnlyList<WordTextMatch> Find(Paragraph paragraph, string text, bool matchCase)
    {
        ArgumentException.ThrowIfNullOrEmpty(text);
        List<Segment> segments = Segments(paragraph.Element);
        string full = string.Concat(segments.Select(segment => segment.Text));
        StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        List<WordTextMatch> matches = [];
        int start = 0;
        while (start <= full.Length - text.Length)
        {
            int index = full.IndexOf(text, start, comparison);
            if (index < 0)
            {
                break;
            }

            matches.Add(new WordTextMatch(paragraph, index, text.Length, full.Substring(index, text.Length)));
            start = index + text.Length;
        }

        return matches;
    }

    internal static int Replace(Paragraph paragraph, string oldText, string newText, bool matchCase)
    {
        ArgumentException.ThrowIfNullOrEmpty(oldText);
        ArgumentNullException.ThrowIfNull(newText);
        PlainText.EnsureValid(newText, nameof(newText));

        List<Segment> segments = Segments(paragraph.Element);
        string full = string.Concat(segments.Select(segment => segment.Text));
        StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        List<(int Start, int Length)> matches = [];
        int start = 0;
        while (start <= full.Length - oldText.Length)
        {
            int index = full.IndexOf(oldText, start, comparison);
            if (index < 0)
            {
                break;
            }

            matches.Add((index, oldText.Length));
            start = index + oldText.Length;
        }

        // The matches are applied from the last to the first so that the offsets of the pending
        // matches stay valid.
        for (int position = matches.Count - 1; position >= 0; position--)
        {
            (int matchStart, int matchLength) = matches[position];
            Apply(segments, matchStart, matchLength, newText);
        }

        return matches.Count;
    }

    private static void Apply(List<Segment> segments, int matchStart, int matchLength, string newText)
    {
        int matchEnd = matchStart + matchLength;
        int first = -1;
        int last = -1;
        for (int index = 0; index < segments.Count; index++)
        {
            Segment segment = segments[index];
            if (segment.Start + segment.Text.Length <= matchStart)
            {
                continue;
            }

            if (segment.Start >= matchEnd)
            {
                break;
            }

            if (first < 0)
            {
                first = index;
            }

            last = index;
        }

        if (first < 0)
        {
            return;
        }

        Segment firstSegment = segments[first];
        int firstLocalStart = matchStart - firstSegment.Start;
        int firstLocalEnd = Math.Min(firstSegment.Text.Length, matchEnd - firstSegment.Start);
        string firstText = firstSegment.Text[..firstLocalStart] + newText;

        Segment lastSegment = segments[last];
        int lastLocalEnd = matchEnd - lastSegment.Start;
        if (first == last)
        {
            firstText += lastSegment.Text[lastLocalEnd..];
        }

        SetText(firstSegment, firstText);

        for (int index = first + 1; index <= last; index++)
        {
            Segment segment = segments[index];
            int localEnd = index == last ? matchEnd - segment.Start : segment.Text.Length;
            string remaining = segment.Text[localEnd..];
            SetText(segment, remaining);
        }
    }

    private static void SetText(Segment segment, string value)
    {
        if (value.Length == 0)
        {
            segment.Element.Remove();
            return;
        }

        segment.Element.Text = value;
        segment.Element.Space = SpaceProcessingModeValues.Preserve;
    }

    private static List<Segment> Segments(W.Paragraph element)
    {
        List<Segment> segments = [];
        int offset = 0;
        foreach (W.Run run in element.Elements<W.Run>())
        {
            offset = Add(segments, run, offset);
        }

        foreach (W.Hyperlink link in element.Elements<W.Hyperlink>())
        {
            foreach (W.Run run in link.Elements<W.Run>())
            {
                offset = Add(segments, run, offset);
            }
        }

        return segments;
    }

    private static int Add(List<Segment> segments, W.Run run, int offset)
    {
        foreach (W.Text text in run.Elements<W.Text>())
        {
            segments.Add(new Segment(text, text.Text, offset));
            offset += text.Text.Length;
        }

        return offset;
    }

    private sealed class Segment
    {
        internal Segment(W.Text element, string text, int start)
        {
            Element = element;
            Text = text;
            Start = start;
        }

        internal W.Text Element { get; }

        internal string Text { get; }

        internal int Start { get; }
    }
}
