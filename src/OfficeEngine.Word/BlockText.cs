using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// Reads the visible text of a sequence of block-level elements (a body, a header or a footer).
/// </summary>
/// <remarks>
/// Each paragraph contributes one line and each table contributes one line per row, which is the same
/// shape <see cref="WordDocument.Text"/> documents. Only the content this version models is read;
/// anything else is skipped here and preserved when the document is saved.
/// </remarks>
internal static class BlockText
{
    internal static string Of(IEnumerable<OpenXmlElement> blocks)
    {
        List<string> lines = [];
        foreach (OpenXmlElement child in blocks)
        {
            if (child is W.Paragraph paragraph)
            {
                lines.Add(new Paragraph(paragraph).Text);
            }
            else if (child is W.Table table)
            {
                lines.Add(OfTable(table));
            }
        }

        return string.Join("\n", lines);
    }

    private static string OfTable(W.Table table)
        => string.Join(
            "\n",
            table.Elements<W.TableRow>().Select(row => string.Join(
                "\t",
                row.Elements<W.TableCell>().Select(cell => new TableCell(cell).Text))));
}
