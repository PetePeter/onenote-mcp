using System.Text;
using OneNoteMcp.Model;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free unit tests for <see cref="BinaryExtractor"/>: the pure XML → decode →
/// extension/filename logic behind the file-extraction tool. Fed the same 1x1 PNG
/// the fixture embeds, so decoded bytes are asserted against real PNG magic.
/// </summary>
public class BinaryExtractorTests
{
    // Clean-room 1x1 transparent PNG, identical to the fixture's embedded image.
    private const string OnePixelPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

    private static readonly byte[] PngMagic = { 0x89, 0x50, 0x4E, 0x47 };

    private static string ImagePageXml(string format = "png", string? size = null) =>
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{page}\">" +
        "<one:Outline><one:OEChildren><one:OE>" +
        $"<one:Image format=\"{format}\">" +
        (size ?? "") +
        $"<one:Data>{OnePixelPngBase64}</one:Data></one:Image>" +
        "</one:OE></one:OEChildren></one:Outline></one:Page>";

    [Fact]
    public void ExtractInlineBinaries_DecodesImageToPngBytes()
    {
        var binaries = BinaryExtractor.ExtractInlineBinaries(ImagePageXml());

        var img = Assert.Single(binaries);
        Assert.Equal("image", img.Type);
        Assert.Equal("png", img.Extension);
        Assert.Equal(PngMagic, img.Bytes[..4]);
        Assert.Equal(Convert.FromBase64String(OnePixelPngBase64).Length, img.Bytes.Length);
    }

    [Fact]
    public void ExtractInlineBinaries_ToleratesWhitespaceInBase64()
    {
        var withNewlines = ImagePageXml().Replace(
            OnePixelPngBase64, "\n  " + OnePixelPngBase64 + "  \n");

        var img = Assert.Single(BinaryExtractor.ExtractInlineBinaries(withNewlines));

        Assert.Equal(PngMagic, img.Bytes[..4]);
    }

    [Fact]
    public void ExtractInlineBinaries_ReadsSizeAsWidthHeight()
    {
        var xml = ImagePageXml(size: "<one:Size width=\"48\" height=\"64\" />");

        var img = Assert.Single(BinaryExtractor.ExtractInlineBinaries(xml));

        Assert.Equal(48, img.Width);
        Assert.Equal(64, img.Height);
    }

    [Fact]
    public void ExtractInlineBinaries_NoSize_LeavesDimensionsNull()
    {
        var img = Assert.Single(BinaryExtractor.ExtractInlineBinaries(ImagePageXml()));

        Assert.Null(img.Width);
        Assert.Null(img.Height);
    }

    [Fact]
    public void ExtractInlineBinaries_EmptyPage_ReturnsEmpty()
    {
        var xml = "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"p\">" +
                  "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[hi]]></one:T></one:OE>" +
                  "</one:OEChildren></one:Outline></one:Page>";

        Assert.Empty(BinaryExtractor.ExtractInlineBinaries(xml));
    }

    [Fact]
    public void InferExtension_PrefersFormatAttribute()
    {
        Assert.Equal("jpg", BinaryExtractor.InferExtension("jpg", new byte[] { 0, 1, 2 }));
    }

    [Fact]
    public void InferExtension_SniffsPngMagicWhenFormatMissing()
    {
        var png = Convert.FromBase64String(OnePixelPngBase64);

        Assert.Equal("png", BinaryExtractor.InferExtension(null, png));
    }

    [Fact]
    public void InferExtension_UnknownFallsBackToBin()
    {
        Assert.Equal("bin", BinaryExtractor.InferExtension("", new byte[] { 0x00, 0x01 }));
    }

    [Fact]
    public void BuildFileName_UsesOneBasedIndexAndExtension()
    {
        Assert.Equal("My Page_1.png", BinaryExtractor.BuildFileName("My Page", 1, "png"));
    }

    [Fact]
    public void BuildFileName_NeutralizesPathTraversalInExtension()
    {
        // A crafted OneNote format attribute must not inject path separators.
        var name = BinaryExtractor.BuildFileName("page", 1, "../../evil");

        Assert.DoesNotContain('/', name);
        Assert.DoesNotContain('\\', name);
    }

    [Fact]
    public void BuildFileName_SanitizesInvalidCharacters()
    {
        var name = BinaryExtractor.BuildFileName("a/b:c*d", 2, "png");

        Assert.DoesNotContain('/', name);
        Assert.DoesNotContain(':', name);
        Assert.DoesNotContain('*', name);
        Assert.EndsWith("_2.png", name);
    }
}
