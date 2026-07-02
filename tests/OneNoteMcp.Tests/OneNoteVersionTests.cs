using OneNoteMcp.Interop;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>Unit tests for OneNoteVersion — pure, no registry access required.</summary>
public sealed class OneNoteVersionTests
{
    [Fact]
    public void ParseCurVer_Office16_ReturnsMajor16With2016Display()
    {
        var version = OneNoteVersion.ParseCurVer("OneNote.Application.16");

        Assert.NotNull(version);
        Assert.Equal(16, version!.Major);
        Assert.Contains("2016", version.DisplayName);
    }

    [Fact]
    public void ParseCurVer_Office15_ReturnsMajor15With2013Display()
    {
        var version = OneNoteVersion.ParseCurVer("OneNote.Application.15");

        Assert.NotNull(version);
        Assert.Equal(15, version!.Major);
        Assert.Equal("2013", version.DisplayName);
    }

    [Fact]
    public void ParseCurVer_Office14_ReturnsMajor14With2010Display()
    {
        var version = OneNoteVersion.ParseCurVer("OneNote.Application.14");

        Assert.NotNull(version);
        Assert.Equal(14, version!.Major);
        Assert.Equal("2010", version.DisplayName);
    }

    [Fact]
    public void ParseCurVer_Office12_ReturnsMajor12With2007Display()
    {
        var version = OneNoteVersion.ParseCurVer("OneNote.Application.12");

        Assert.NotNull(version);
        Assert.Equal(12, version!.Major);
        Assert.Equal("2007", version.DisplayName);
    }

    [Fact]
    public void ParseCurVer_UnknownMajor_ReturnsOfficePrefix()
    {
        var version = OneNoteVersion.ParseCurVer("OneNote.Application.99");

        Assert.NotNull(version);
        Assert.Equal(99, version!.Major);
        Assert.Contains("99", version.DisplayName);
    }

    [Fact]
    public void ParseCurVer_Null_ReturnsNull()
    {
        Assert.Null(OneNoteVersion.ParseCurVer(null));
    }

    [Fact]
    public void ParseCurVer_EmptyString_ReturnsNull()
    {
        Assert.Null(OneNoteVersion.ParseCurVer(string.Empty));
    }

    [Fact]
    public void ParseCurVer_GarbageString_ReturnsNull()
    {
        Assert.Null(OneNoteVersion.ParseCurVer("not-a-version"));
    }

    [Fact]
    public void ParseCurVer_NoTrailingNumber_ReturnsNull()
    {
        Assert.Null(OneNoteVersion.ParseCurVer("OneNote.Application."));
    }
}
