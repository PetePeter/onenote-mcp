using System.Text.Json;
using OneNoteMcp.Model;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// End-to-end COM integration tests for the page read tools. They pull real page
/// content from the shared <see cref="FixtureNotebook"/> and assert against its
/// known text, image data, and title. They early-return (skip) when the fixture
/// could not be built.
/// </summary>
[Collection("OneNote COM")]
public sealed class PageReadToolsIntegrationTests
{
    private static readonly JsonSerializerOptions JsonOpts =
        new() { PropertyNameCaseInsensitive = true };

    private readonly FixtureNotebook _fx;

    public PageReadToolsIntegrationTests(FixtureNotebook fx) => _fx = fx;

    [Fact]
    public void GetPage_Basic_ReturnsXmlContainingKnownText()
    {
        if (!_fx.Available) return;

        var xml = PageReadTools.GetPage(_fx.TextPageId, "basic");

        Assert.Contains(_fx.KnownText, xml);
    }

    [Fact]
    public void GetPage_All_IncludesImageData()
    {
        if (!_fx.Available) return;

        var xml = PageReadTools.GetPage(_fx.ImagePageId, "all");

        Assert.Contains("Image", xml);
        Assert.Contains("Data", xml);
    }

    [Fact]
    public void GetPageInfo_ReturnsFixtureTitle()
    {
        if (!_fx.Available) return;

        var json = PageReadTools.GetPageInfo(_fx.TextPageId);
        var meta = JsonSerializer.Deserialize<PageMetadata>(json, JsonOpts)!;

        Assert.Equal(_fx.KnownTitle, meta.Title);
        Assert.Equal(_fx.TextPageId, meta.Id);
    }
}
