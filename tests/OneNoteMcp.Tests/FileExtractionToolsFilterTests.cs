using OneNoteMcp.Model;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Unit tests for <see cref="FileExtractionTools.ApplyFilter"/>: verifies the
/// caller-facing selection contract, in particular that ink is its own category
/// (selected by "ink", excluded from "files", included in "all").
/// </summary>
public class FileExtractionToolsFilterTests
{
    private static InlineBinary Bin(string type) =>
        new(new byte[] { 0, 1, 2, 3 }, type, type == "ink" ? "isf" : "png", null, null, "", null);

    private static readonly IReadOnlyList<InlineBinary> Sample =
        new[] { Bin("image"), Bin("file"), Bin("ink") };

    [Fact]
    public void ApplyFilter_Ink_SelectsOnlyInk()
    {
        var result = FileExtractionTools.ApplyFilter(Sample, "ink");

        var only = Assert.Single(result);
        Assert.Equal("ink", only.Type);
    }

    [Fact]
    public void ApplyFilter_Files_ExcludesInk()
    {
        var result = FileExtractionTools.ApplyFilter(Sample, "files");

        var only = Assert.Single(result);
        Assert.Equal("file", only.Type);
    }

    [Fact]
    public void ApplyFilter_All_IncludesInk()
    {
        var result = FileExtractionTools.ApplyFilter(Sample, "all").ToList();

        Assert.Equal(3, result.Count);
        Assert.Contains(result, b => b.Type == "ink");
    }

    [Fact]
    public void ApplyFilter_Images_Unaffected()
    {
        var result = FileExtractionTools.ApplyFilter(Sample, "images");

        var only = Assert.Single(result);
        Assert.Equal("image", only.Type);
    }
}
