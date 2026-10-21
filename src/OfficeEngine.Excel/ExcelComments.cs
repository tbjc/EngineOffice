using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>Builds and reads the comments part of a worksheet and the legacy drawing that goes with it.</summary>
internal static class ExcelComments
{
    internal const string DefaultAuthor = "OfficeEngine";

    internal static S.Comments Build(
        IReadOnlyList<(string Reference, string Text, string Author)> comments,
        out IReadOnlyList<string> authors)
    {
        List<string> authorList = [];
        foreach ((string _, string _, string author) in comments)
        {
            if (!authorList.Contains(author, StringComparer.Ordinal))
            {
                authorList.Add(author);
            }
        }

        authors = authorList;

        S.Authors authorsElement = new();
        foreach (string author in authorList)
        {
            authorsElement.AppendChild(new S.Author(author));
        }

        S.CommentList commentList = new();
        foreach ((string reference, string text, string author) in comments)
        {
            S.Run run = new(
                new S.RunProperties(
                    new S.FontSize { Val = 9d },
                    new S.Color { Indexed = 81U },
                    new S.RunFont { Val = "Tahoma" },
                    new S.FontFamily { Val = 2 }),
                new S.Text(text) { Space = SpaceProcessingModeValues.Preserve });
            commentList.AppendChild(new S.Comment
            {
                Reference = reference,
                AuthorId = (uint)authorList.IndexOf(author),
                CommentText = new S.CommentText(run),
            });
        }

        return new S.Comments(authorsElement, commentList);
    }

    /// <summary>
    /// Rebuilds the legacy drawing that holds the shape of every note. Excel needs it to show the
    /// markers of the notes.
    /// </summary>
    internal static string BuildDrawing(IReadOnlyList<(string Reference, string Text, string Author)> comments)
    {
        StringBuilder builder = new();
        builder.Append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" ");
        builder.Append("xmlns:o=\"urn:schemas-microsoft-com:office:office\" ");
        builder.Append("xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
        builder.Append("<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>");
        builder.Append("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" ");
        builder.Append("path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/>");
        builder.Append("<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>");

        for (int index = 0; index < comments.Count; index++)
        {
            (string reference, string _, string _) = comments[index];
            int column = 0;
            int row = 0;
            if (!CellReference.TryParse(reference, out column, out row))
            {
                column = 0;
                row = 0;
            }
            column--;
            row--;

            builder.Append(CultureInfo.InvariantCulture, $"<v:shape id=\"_x0000_s{1025 + index}\" ");
            builder.Append("type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;");
            builder.Append("margin-top:5pt;width:200pt;height:100pt;z-index:");
            builder.Append(CultureInfo.InvariantCulture, $"{index + 1}");
            builder.Append(";visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">");
            builder.Append("<v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
            builder.Append("<v:path o:connecttype=\"none\"/>");
            builder.Append("<v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox>");
            builder.Append("<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/>");
            builder.Append(CultureInfo.InvariantCulture, $"<x:Anchor>{column + 1}, 15, {row}, 2, {column + 3}, 31, {row + 4}, 1</x:Anchor>");
            builder.Append("<x:AutoFill>False</x:AutoFill>");
            builder.Append(CultureInfo.InvariantCulture, $"<x:Row>{row}</x:Row><x:Column>{column}</x:Column>");
            builder.Append("</x:ClientData></v:shape>");
        }

        builder.Append("</xml>");
        return builder.ToString();
    }

    internal static List<(string Reference, string Text, string Author)> Read(S.Comments comments)
    {
        List<string> authors = [];
        if (comments.Authors is not null)
        {
            foreach (S.Author author in comments.Authors.Elements<S.Author>())
            {
                authors.Add(author.Text);
            }
        }

        List<(string Reference, string Text, string Author)> result = [];
        if (comments.CommentList is null)
        {
            return result;
        }

        foreach (S.Comment comment in comments.CommentList.Elements<S.Comment>())
        {
            uint authorId = comment.AuthorId?.Value ?? 0U;
            string author = authorId < authors.Count ? authors[(int)authorId] : string.Empty;
            result.Add((
                comment.Reference?.Value ?? string.Empty,
                comment.CommentText?.InnerText ?? string.Empty,
                author));
        }

        return result;
    }
}
