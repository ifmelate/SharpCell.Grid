using System.Runtime.CompilerServices;
using SharpCell.Grid.Docs;

namespace SharpCell.Grid.Tests.Docs;

public class DocsTests
{
    private static RepositoryLayout Repository([CallerFilePath] string here = "") =>
        RepositoryLayout.At(Path.Combine(Path.GetDirectoryName(here)!, "..", "..", ".."));

    [Fact]
    public void The_committed_documentation_matches_what_the_tool_generates()
    {
        var layout = Repository();
        var differences = Generator.Differences(layout, Generator.Run(layout));
        Assert.True(differences.Count == 0,
            "Run 'dotnet run --project tools/SharpCell.Grid.Docs -f net10.0 -- generate':\n" + string.Join("\n", differences));
    }

    [Fact]
    public void Every_public_member_has_a_summary()
    {
        var assemblies = ApiReference.DocumentedAssemblies(Repository());
        var docs = XmlDocs.Load(assemblies);
        var empty = assemblies
            .SelectMany(a => a.GetExportedTypes())
            .SelectMany(t => ApiReference.MembersOf(t).Prepend(t))
            .Select(DocId.Of)
            .Where(id => !docs.Has(id) || docs.Summary(id, _ => null).Length == 0)
            .ToList();

        Assert.True(empty.Count == 0, "No summary for:\n" + string.Join("\n", empty));
    }

    [Fact]
    public void The_site_has_a_page_for_every_listed_page_with_its_pictures_and_llms_txt()
    {
        var output = Path.Combine(Path.GetTempPath(), "sharpcell-grid-site-" + Guid.NewGuid().ToString("N"));
        try
        {
            var layout = Repository();
            SiteBuilder.Build(layout, output);

            var config = SiteConfig.Load(layout.Full("docs/site.json"));
            foreach (var page in config.Sections.SelectMany(s => s.Pages))
            {
                var html = Path.Combine(output, Path.ChangeExtension(page.Path, ".html"));
                Assert.True(File.Exists(html), $"{page.Path} has no HTML page");
                Assert.Contains("· SharpCell.Grid</title>", File.ReadAllText(html));
            }

            Assert.True(File.Exists(Path.Combine(output, "api", "index.html")));
            Assert.True(File.Exists(Path.Combine(output, "llms.txt")));
            foreach (var image in Directory.EnumerateFiles(layout.Full("docs/images")))
                Assert.True(File.Exists(Path.Combine(output, "images", Path.GetFileName(image))));
            Assert.Contains("images/styles.png", File.ReadAllText(Path.Combine(output, "index.html")));
        }
        finally
        {
            if (Directory.Exists(output))
                Directory.Delete(output, recursive: true);
        }
    }
}
