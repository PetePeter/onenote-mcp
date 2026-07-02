using OneNoteMcp.Model;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free unit tests for <see cref="PageXmlValidator"/>. They pin the guard that
/// runs before any UpdatePageContent COM call: the XML must be well-formed and
/// rooted in a OneNote &lt;Page&gt; element in a onenote schema namespace, with an
/// actionable message on every rejection.
/// </summary>
public class PageXmlValidatorTests
{
    private const string OneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private static readonly string ValidPage =
        $"<one:Page xmlns:one=\"{OneNs}\" ID=\"{{ID}}\">" +
        "<one:Title><one:OE><one:T><![CDATA[Hi]]></one:T></one:OE></one:Title>" +
        "</one:Page>";

    [Fact]
    public void Accepts_a_valid_2013_page()
    {
        // Should not throw.
        PageXmlValidator.Validate(ValidPage);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void Rejects_empty_input(string? xml)
    {
        var ex = Assert.Throws<ArgumentException>(() => PageXmlValidator.Validate(xml!));
        Assert.Contains("empty", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rejects_malformed_xml_with_actionable_message()
    {
        // Unclosed element — not well-formed.
        var ex = Assert.Throws<ArgumentException>(
            () => PageXmlValidator.Validate($"<one:Page xmlns:one=\"{OneNs}\"><one:Title>"));
        Assert.Contains("well-formed", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rejects_foreign_namespace_root()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => PageXmlValidator.Validate("<Page xmlns=\"http://example.com/not-onenote\" ID=\"x\" />"));
        Assert.Contains("namespace", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rejects_no_namespace_root()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => PageXmlValidator.Validate("<Page ID=\"x\" />"));
        Assert.Contains("namespace", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rejects_wrong_root_element_even_in_onenote_namespace()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => PageXmlValidator.Validate($"<one:Section xmlns:one=\"{OneNs}\" ID=\"x\" />"));
        Assert.Contains("Page", ex.Message);
    }

    [Fact]
    public void Rejects_page_missing_id_attribute()
    {
        // A well-formed OneNote <Page> with no ID would silently create a NEW page
        // on UpdatePageContent instead of updating the intended one.
        var ex = Assert.Throws<ArgumentException>(
            () => PageXmlValidator.Validate($"<one:Page xmlns:one=\"{OneNs}\" />"));
        Assert.Contains("ID", ex.Message);
    }

    [Fact]
    public void Rejects_wrong_case_root_element()
    {
        // XML local names are case-sensitive; lowercase <one:page> is not <one:Page>.
        var ex = Assert.Throws<ArgumentException>(
            () => PageXmlValidator.Validate($"<one:page xmlns:one=\"{OneNs}\" ID=\"x\" />"));
        Assert.Contains("Page", ex.Message);
    }
}
