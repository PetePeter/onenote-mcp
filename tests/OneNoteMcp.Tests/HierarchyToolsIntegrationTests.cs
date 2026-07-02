using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// End-to-end COM integration tests for the hierarchy read tools. They exercise
/// the real OneNote application through the shared <see cref="FixtureNotebook"/>
/// and assert the fixture's notebook/section/pages surface by their object IDs.
/// They early-return (skip) when the fixture could not be built.
/// </summary>
[Collection("OneNote COM")]
public sealed class HierarchyToolsIntegrationTests
{
    private static readonly JsonSerializerOptions JsonOpts =
        new() { PropertyNameCaseInsensitive = true };

    private readonly FixtureNotebook _fx;

    public HierarchyToolsIntegrationTests(FixtureNotebook fx) => _fx = fx;

    [Fact]
    public void ListNotebooks_IncludesFixtureNotebook()
    {
        if (!_fx.Available) return;

        var json = HierarchyTools.ListNotebooks();
        var notebooks = JsonSerializer.Deserialize<NotebookInfo[]>(json, JsonOpts)!;

        Assert.Contains(notebooks, n => n.Id == _fx.NotebookId);
    }

    [Fact]
    public void GetHierarchy_Pages_ReturnsFixturePagesById()
    {
        if (!_fx.Available) return;

        var xml = HierarchyTools.GetHierarchy(_fx.NotebookId, "pages");

        Assert.Contains(_fx.TextPageId, xml);
        Assert.Contains(_fx.ImagePageId, xml);
    }

    [Fact]
    public void FindPages_MatchesFixturePageTitle()
    {
        if (!_fx.Available) return;

        string json;
        try
        {
            json = HierarchyTools.FindPages(_fx.KnownTitle);
        }
        catch (COMException)
        {
            // FindPages requires the search index; on some environments it is not
            // available. Treat as a skip rather than a failure.
            return;
        }

        var pages = JsonSerializer.Deserialize<PageMatch[]>(json, JsonOpts)!;

        // Both fixture pages carry KnownTitle; at least one fixture page must match.
        var fixturePageIds = new[] { _fx.TextPageId, _fx.ImagePageId };
        Assert.Contains(pages, p => fixturePageIds.Contains(p.Id));
    }
}
