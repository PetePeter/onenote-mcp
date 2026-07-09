using System.IO;
using System.Xml.Linq;
using OneNoteMcp.Model;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Pure unit tests (no COM) for <see cref="EmbedXmlBuilder"/>: the type-routing
/// (image inline vs file attachment), namespace selection, base64 round-trip, and
/// optional outline placement that back the onenote_embed_file tool.
/// </summary>
public sealed class EmbedXmlBuilderTests
{
    private const string Ns2007 = "http://schemas.microsoft.com/office/onenote/2007/onenote";
    private const string Ns2013 = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    // Minimal 1x1 PNG.
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    [Fact]
    public void ImageFile_IsInlinedAsBase64Image()
    {
        var r = EmbedXmlBuilder.Build("{P}", @"C:\pics\logo.png", PngBytes, null, 16, null, null);

        Assert.Equal("image", r.Kind);
        Assert.Equal("png", r.Format);

        XNamespace one = Ns2013;
        var img = XDocument.Parse(r.Xml).Descendants(one + "Image").Single();
        Assert.Equal("png", img.Attribute("format")!.Value);
        var data = img.Element(one + "Data")!.Value;
        Assert.Equal(PngBytes, Convert.FromBase64String(data));
        Assert.Empty(XDocument.Parse(r.Xml).Descendants(one + "InsertedFile"));
    }

    [Fact]
    public void NonImageFile_IsAttachedByPathNotBase64()
    {
        var r = EmbedXmlBuilder.Build("{P}", @"C:\docs\report.pdf", new byte[] { 1, 2, 3 }, null, 16, null, null);

        Assert.Equal("file", r.Kind);
        XNamespace one = Ns2013;
        var inserted = XDocument.Parse(r.Xml).Descendants(one + "InsertedFile").Single();
        Assert.Equal(Path.GetFullPath(@"C:\docs\report.pdf"), inserted.Attribute("pathCache")!.Value);
        Assert.Equal("report.pdf", inserted.Attribute("preferredName")!.Value);
        // No base64 data element for attachments.
        Assert.Empty(XDocument.Parse(r.Xml).Descendants(one + "Data"));
    }

    [Fact]
    public void PreferredName_OverridesFileName()
    {
        var r = EmbedXmlBuilder.Build("{P}", @"C:\docs\tmp1234.bin", new byte[] { 9 }, "Quarterly.xlsx", 16, null, null);

        XNamespace one = Ns2013;
        var inserted = XDocument.Parse(r.Xml).Descendants(one + "InsertedFile").Single();
        Assert.Equal("Quarterly.xlsx", inserted.Attribute("preferredName")!.Value);
    }

    [Fact]
    public void Major12_UsesThe2007Namespace()
    {
        var r = EmbedXmlBuilder.Build("{P}", @"C:\pics\logo.png", PngBytes, null, 12, null, null);

        var root = XDocument.Parse(r.Xml).Root!;
        Assert.Equal(Ns2007, root.Name.NamespaceName);
    }

    [Fact]
    public void PageId_IsCarriedOnTheRoot()
    {
        var r = EmbedXmlBuilder.Build("{PAGE-42}", @"C:\pics\logo.png", PngBytes, null, 16, null, null);
        Assert.Equal("{PAGE-42}", XDocument.Parse(r.Xml).Root!.Attribute("ID")!.Value);
    }

    [Fact]
    public void Position_IsEmittedWhenBothCoordsGiven()
    {
        var r = EmbedXmlBuilder.Build("{P}", @"C:\pics\logo.png", PngBytes, null, 16, 72.0, 144.5);

        XNamespace one = Ns2013;
        var pos = XDocument.Parse(r.Xml).Descendants(one + "Position").Single();
        Assert.Equal("72", pos.Attribute("x")!.Value);
        Assert.Equal("144.5", pos.Attribute("y")!.Value);
    }

    [Fact]
    public void Position_IsOmittedWhenEitherCoordMissing()
    {
        var r = EmbedXmlBuilder.Build("{P}", @"C:\pics\logo.png", PngBytes, null, 16, 72.0, null);

        XNamespace one = Ns2013;
        Assert.Empty(XDocument.Parse(r.Xml).Descendants(one + "Position"));
    }

    [Fact]
    public void ExtensionlessFile_TypedByMagicSniff()
    {
        // No extension, but PNG magic bytes -> image.
        var r = EmbedXmlBuilder.Build("{P}", @"C:\tmp\blob", PngBytes, null, 16, null, null);
        Assert.Equal("image", r.Kind);
        Assert.Equal("png", r.Format);
    }

    [Fact]
    public void ResolveOutputDir_ReturnsCallerDirWhenGiven()
    {
        Assert.Equal(@"C:\out", BinaryExtractor.ResolveOutputDir(@"C:\out"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ResolveOutputDir_CreatesTempDirWhenMissing(string? given)
    {
        var dir = BinaryExtractor.ResolveOutputDir(given);
        Assert.True(Directory.Exists(dir));
        Assert.StartsWith(Path.GetTempPath(), dir);
        Directory.Delete(dir);
    }
}
