using System;
using System.Linq;
using OneNoteMcp.Model;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free unit tests for the static per-major capability table. Encodes the
/// P-0545 spike result: v12 (2007) supports 17 of the 23 wrapped IApplication
/// methods; the 6 IApplication2-4 additions are absent. Modern majors (14/15/16)
/// support all 23.
/// </summary>
public sealed class OneNoteCapabilitiesTests
{
    // The 6 methods added in 2010+ (IApplication2-4), absent on v12.
    public static readonly Capability[] ModernOnly =
    [
        Capability.NavigateToUrl,
        Capability.GetWebHyperlinkToObject,
        Capability.MergeFiles,
        Capability.MergeSections,
        Capability.SyncHierarchy,
        Capability.SetFilingLocation,
    ];

    [Fact]
    public void Supports_SyncHierarchy_TrueOn16_FalseOn12()
    {
        Assert.True(OneNoteCapabilities.Supports(16, Capability.SyncHierarchy));
        Assert.False(OneNoteCapabilities.Supports(12, Capability.SyncHierarchy));
    }

    [Theory]
    [MemberData(nameof(ModernOnlyData))]
    public void Supports_ModernOnlyMethods_FalseOn12(Capability cap)
        => Assert.False(OneNoteCapabilities.Supports(12, cap));

    public static IEnumerable<object[]> ModernOnlyData => ModernOnly.Select(c => new object[] { c });

    [Theory]
    [InlineData(Capability.GetHierarchy)]
    [InlineData(Capability.FindPages)]
    [InlineData(Capability.GetBinaryPageContent)]
    [InlineData(Capability.FindMeta)]
    [InlineData(Capability.NavigateTo)]
    [InlineData(Capability.GetHyperlinkToObject)]
    public void Supports_SharedMethods_TrueOn12(Capability cap)
        => Assert.True(OneNoteCapabilities.Supports(12, cap));

    [Theory]
    [InlineData(12)]
    [InlineData(14)]
    [InlineData(15)]
    [InlineData(16)]
    public void Supports_EveryCapabilityDecidedForKnownMajor(int major)
    {
        // Table completeness: every Capability has an explicit decision for each
        // known major — no throw for any enum value.
        foreach (Capability cap in Enum.GetValues<Capability>())
            _ = OneNoteCapabilities.Supports(major, cap);
    }

    [Fact]
    public void Supports_ModernMajors_AllTrue()
    {
        foreach (int major in new[] { 14, 15, 16 })
            foreach (Capability cap in Enum.GetValues<Capability>())
                Assert.True(OneNoteCapabilities.Supports(major, cap), $"{major}:{cap}");
    }

    [Fact]
    public void Supports_V12_HasExactly17Supported()
    {
        int count = Enum.GetValues<Capability>().Count(c => OneNoteCapabilities.Supports(12, c));
        Assert.Equal(17, count);
    }

    [Fact]
    public void Supports_UnknownMajor_Throws()
        => Assert.Throws<ArgumentException>(() => OneNoteCapabilities.Supports(99, Capability.GetHierarchy));

    [Fact]
    public void SupportedBy_V12_Returns17()
        => Assert.Equal(17, OneNoteCapabilities.SupportedBy(12).Count);

    [Fact]
    public void SupportedBy_V16_ReturnsAll23()
        => Assert.Equal(Enum.GetValues<Capability>().Length, OneNoteCapabilities.SupportedBy(16).Count);

    [Fact]
    public void Capability_HasExactly23Values()
        => Assert.Equal(23, Enum.GetValues<Capability>().Length);
}
