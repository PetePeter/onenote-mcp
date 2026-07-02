using OneNoteMcp.Model;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free unit tests for <see cref="VersionResolver"/> — friendly version token
/// (year, major int, or raw CLSID) → canonical CLSID string. There is NO default:
/// unknown or empty tokens throw (see P-0545 DECISIONS — every tool passes an explicit
/// version).
/// </summary>
public sealed class VersionResolverTests
{
    const string Clsid2007 = "{0039FFEC-A022-4232-8274-6B34787BFC27}";
    const string Clsid2010 = "{D7FAC39E-7FF1-49AA-98CF-A1DDD316337E}";
    const string ClsidModern = "{DC67E480-C3CB-49F8-8232-60B0C2056C8E}";

    [Theory]
    [InlineData("2007", Clsid2007)]
    [InlineData("2010", Clsid2010)]
    [InlineData("2013", ClsidModern)] // classic desktop CLSID, shared with 2016
    [InlineData("2016", ClsidModern)]
    public void Resolve_MapsFriendlyYearTokens(string token, string expected)
        => Assert.Equal(expected, VersionResolver.Resolve(token));

    [Theory]
    [InlineData("12", Clsid2007)]
    [InlineData("14", Clsid2010)]
    [InlineData("16", ClsidModern)]
    public void Resolve_MapsMajorIntTokens(string token, string expected)
        => Assert.Equal(expected, VersionResolver.Resolve(token));

    [Fact]
    public void Resolve_RawClsidNoBraces_ReturnsNormalizedUppercaseBraces()
        => Assert.Equal(Clsid2007, VersionResolver.Resolve("0039ffec-a022-4232-8274-6b34787bfc27"));

    [Fact]
    public void Resolve_RawClsidWithBraces_Normalizes()
        => Assert.Equal(ClsidModern, VersionResolver.Resolve("{dc67e480-c3cb-49f8-8232-60b0c2056c8e}"));

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void Resolve_EmptyOrWhitespace_Throws(string? token)
        => Assert.Throws<ArgumentException>(() => VersionResolver.Resolve(token!));

    [Theory]
    [InlineData("2099")]
    [InlineData("banana")]
    [InlineData("99")]
    public void Resolve_UnknownToken_Throws(string token)
        => Assert.Throws<ArgumentException>(() => VersionResolver.Resolve(token));
}
