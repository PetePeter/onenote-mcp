using System.Text.Json;
using OneNoteMcp.Model;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// End-to-end COM integration tests for the file-extraction tool. They extract the
/// fixture image page's embedded 1x1 PNG to a scratch directory and assert a real,
/// decodable PNG lands on disk at an absolute path. Each test cleans up its own
/// output directory and early-returns (skips) when the fixture is unavailable.
/// </summary>
[Collection("OneNote COM")]
public sealed class FileExtractionToolsIntegrationTests
{
    private static readonly JsonSerializerOptions JsonOpts =
        new() { PropertyNameCaseInsensitive = true };

    private static readonly byte[] PngMagic = { 0x89, 0x50, 0x4E, 0x47 };

    private readonly FixtureNotebook _fx;

    public FileExtractionToolsIntegrationTests(FixtureNotebook fx) => _fx = fx;

    private static string NewOutputDir() =>
        Path.Combine(Path.GetTempPath(), "onenote-mcp-tests", "extract-" + Guid.NewGuid().ToString("N"));

    [Fact]
    public void ExtractPageFiles_WritesDecodablePngAtAbsolutePath()
    {
        if (!_fx.Available) return;

        var outDir = NewOutputDir();
        try
        {
            var json = FileExtractionTools.ExtractPageFiles(FixtureNotebook.TestVersion, _fx.ImagePageId, outDir, "images");
            var files = JsonSerializer.Deserialize<List<ExtractedFile>>(json, JsonOpts)!;

            var file = Assert.Single(files);
            Assert.Equal("image", file.Type);
            Assert.True(Path.IsPathRooted(file.Path), "returned path must be absolute");
            Assert.True(File.Exists(file.Path), "extracted file must exist on disk");
            Assert.Equal(PngMagic, File.ReadAllBytes(file.Path)[..4]);
        }
        finally
        {
            if (Directory.Exists(outDir)) Directory.Delete(outDir, recursive: true);
        }
    }

    [Fact]
    public void ExtractPageFiles_CreatesMissingOutputDir()
    {
        if (!_fx.Available) return;

        var outDir = NewOutputDir();
        Assert.False(Directory.Exists(outDir));
        try
        {
            FileExtractionTools.ExtractPageFiles(FixtureNotebook.TestVersion, _fx.ImagePageId, outDir, "all");

            Assert.True(Directory.Exists(outDir));
        }
        finally
        {
            if (Directory.Exists(outDir)) Directory.Delete(outDir, recursive: true);
        }
    }

    [Fact]
    public void ExtractPageFiles_PageWithoutImages_ReturnsEmptyList()
    {
        if (!_fx.Available) return;

        var outDir = NewOutputDir();
        try
        {
            var json = FileExtractionTools.ExtractPageFiles(FixtureNotebook.TestVersion, _fx.TextPageId, outDir, "all");
            var files = JsonSerializer.Deserialize<List<ExtractedFile>>(json, JsonOpts)!;

            Assert.Empty(files);
        }
        finally
        {
            if (Directory.Exists(outDir)) Directory.Delete(outDir, recursive: true);
        }
    }
}
