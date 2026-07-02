using System;
using OneNoteMcp.Model;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free unit tests for the PageInfo detail-level mapping and page-metadata
/// parsing used by the page read tools.
/// </summary>
public class PageInfoMapperTests
{
    [Theory]
    [InlineData("basic", 0)]
    [InlineData("selection", 1)]
    [InlineData("binaryData", 2)]
    [InlineData("fileType", 4)]
    [InlineData("all", 7)]
    public void MapDetail_MapsKnownLevelsToPageInfoFlags(string detail, int expected)
    {
        Assert.Equal(expected, PageInfoMapper.MapDetail(detail));
    }

    [Theory]
    [InlineData("BASIC")]
    [InlineData("All")]
    [InlineData("BinaryData")]
    public void MapDetail_IsCaseInsensitive(string detail)
    {
        // Should not throw; case must not matter.
        _ = PageInfoMapper.MapDetail(detail);
    }

    [Fact]
    public void MapDetail_NullOrEmpty_DefaultsToBasic()
    {
        Assert.Equal(0, PageInfoMapper.MapDetail(null!));
        Assert.Equal(0, PageInfoMapper.MapDetail(""));
    }

    [Fact]
    public void MapDetail_UnknownDetail_Throws()
    {
        Assert.Throws<ArgumentException>(() => PageInfoMapper.MapDetail("nonsense"));
    }

    [Fact]
    public void ParseMetadata_ExtractsTitleAndPageAttributes()
    {
        const string ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        var xml =
            $"<one:Page xmlns:one=\"{ns}\" ID=\"{{page-id}}\" name=\"Attr Title\" " +
            "dateTime=\"2026-07-02T09:00:00.000Z\" lastModifiedTime=\"2026-07-02T10:00:00.000Z\" " +
            "pageLevel=\"1\" author=\"Jane Doe\">" +
            "<one:Title><one:OE><one:T><![CDATA[Real Title]]></one:T></one:OE></one:Title>" +
            "</one:Page>";

        var meta = PageInfoMapper.ParseMetadata(xml);

        // Title text run wins over the name attribute.
        Assert.Equal("Real Title", meta.Title);
        Assert.Equal("{page-id}", meta.Id);
        Assert.Equal("2026-07-02T09:00:00.000Z", meta.DateTime);
        Assert.Equal("2026-07-02T10:00:00.000Z", meta.LastModified);
        Assert.Equal("1", meta.Level);
        Assert.Equal("Jane Doe", meta.Author);
    }

    [Fact]
    public void ParseMetadata_FallsBackToNameAttributeWhenNoTitleElement()
    {
        const string ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        var xml = $"<one:Page xmlns:one=\"{ns}\" ID=\"p\" name=\"Only Attr\" pageLevel=\"2\" />";

        var meta = PageInfoMapper.ParseMetadata(xml);

        Assert.Equal("Only Attr", meta.Title);
        Assert.Equal("2", meta.Level);
    }
}
