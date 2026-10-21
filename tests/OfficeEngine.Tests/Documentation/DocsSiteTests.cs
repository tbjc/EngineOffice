using System.Text.RegularExpressions;

namespace OfficeEngine.Tests.Documentation;

/// <summary>
/// The Docusaurus site lives in <c>docs-site</c>. These tests check that its pages and its sidebar
/// stay in sync: a page that is not in the sidebar would be invisible, and a sidebar entry without a
/// page would break the build of the site.
/// </summary>
public sealed class DocsSiteTests
{
    [Fact]
    public void EveryPageHasAnIdThatMatchesItsFileName()
    {
        foreach (string page in Pages())
        {
            string text = File.ReadAllText(PagePath(page));
            string name = Path.GetFileName(page);
            Assert.Matches($@"(?m)^id:\s*{Regex.Escape(name)}\s*$", text);
        }
    }

    [Fact]
    public void TheSidebarListsExactlyThePagesOfTheSite()
    {
        Assert.Equal(Pages().ToHashSet(StringComparer.Ordinal), SidebarEntries().ToHashSet(StringComparer.Ordinal));
    }

    [Fact]
    public void TheSidebarStartsWithTheIndexAndKeepsTheRequestedOrder()
    {
        List<string> entries = [.. SidebarEntries()];

        Assert.Equal("index", entries[0]);
        Assert.True(entries.IndexOf("word/crear-y-guardar") < entries.IndexOf("excel/crear-y-guardar"));
        Assert.True(entries.IndexOf("excel/crear-y-guardar") < entries.IndexOf("pdf/convertir-a-pdf"));
    }

    [Fact]
    public void EveryPageHasAtLeastOneExample()
    {
        foreach (string page in Pages())
        {
            if (string.Equals(page, "index", StringComparison.Ordinal))
            {
                continue;
            }

            Assert.Contains("```csharp", File.ReadAllText(PagePath(page)), StringComparison.Ordinal);
        }
    }

    private static IEnumerable<string> Pages()
    {
        string root = Path.Combine(RepositoryRoot(), "docs-site", "docs");
        foreach (string file in Directory.EnumerateFiles(root, "*.md", SearchOption.AllDirectories))
        {
            yield return Path.ChangeExtension(Path.GetRelativePath(root, file), null).Replace('\\', '/');
        }
    }

    private static IEnumerable<string> SidebarEntries()
    {
        string sidebar = File.ReadAllText(Path.Combine(RepositoryRoot(), "docs-site", "sidebars.ts"));
        foreach (Match match in Regex.Matches(sidebar, "'(?<id>[a-z0-9/-]+)'"))
        {
            string id = match.Groups["id"].Value;
            if (id.Contains('/') || string.Equals(id, "index", StringComparison.Ordinal))
            {
                yield return id;
            }
        }
    }

    private static string PagePath(string page)
        => Path.Combine(RepositoryRoot(), "docs-site", "docs", page.Replace('/', Path.DirectorySeparatorChar) + ".md");

    private static string RepositoryRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeEngine.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("No se encuentra la raíz del repositorio.");
    }
}
