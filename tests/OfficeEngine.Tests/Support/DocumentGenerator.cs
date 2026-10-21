using OfficeEngine.Word;

namespace OfficeEngine.Tests.Support;

/// <summary>
/// Generates documents with a deterministic pseudo-random sequence, so that a round-trip can be
/// checked over many different shapes without depending on a fixture file.
/// </summary>
internal sealed class DocumentGenerator
{
    private uint _state;

    internal DocumentGenerator(uint seed) => _state = seed == 0U ? 1U : seed;

    internal int Next(int exclusiveMaximum)
    {
        _state ^= _state << 13;
        _state ^= _state >> 17;
        _state ^= _state << 5;
        return (int)(_state % (uint)exclusiveMaximum);
    }

    internal string NextText()
    {
        string[] words = ["informe", "ventas", "trimestre", "total", "importe", "cliente", "nota"];
        int count = 1 + Next(6);
        List<string> parts = [];
        for (int index = 0; index < count; index++)
        {
            parts.Add(words[Next(words.Length)]);
        }

        return string.Join(' ', parts);
    }

    internal WordDocument CreateDocument()
    {
        WordDocument document = WordDocument.Create();
        int paragraphs = 3 + Next(8);
        for (int index = 0; index < paragraphs; index++)
        {
            Paragraph paragraph = document.AddParagraph(NextText());
            int runs = Next(3);
            for (int run = 0; run < runs; run++)
            {
                Run created = paragraph.AddRun(" " + NextText());
                created.Bold = Next(2) == 0;
                created.Italic = Next(3) == 0;
                created.FontSizePoints = 9 + Next(6);
            }

            if (Next(4) == 0)
            {
                paragraph.Alignment = ParagraphAlignment.Center;
            }
        }

        if (Next(2) == 0)
        {
            Table table = document.AddTable(1 + Next(2), 1 + Next(3));
            for (int row = 0; row < table.RowCount; row++)
            {
                for (int column = 0; column < table.ColumnCount; column++)
                {
                    table.Rows[row].Cells[column].Text = NextText();
                }
            }
        }

        if (Next(3) == 0)
        {
            document.Sections[0].PageSize = PageSize.A4;
        }

        return document;
    }
}
