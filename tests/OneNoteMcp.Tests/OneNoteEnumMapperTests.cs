using OneNoteMcp.Interop;
using OneNoteMcp.Model;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free unit tests for the friendly-name → OneNote enum mapping used by the
/// location / filing tools. Verifies known names map to the interop's integer
/// values, case-insensitivity, and that unknown names fail loudly.
/// </summary>
public class OneNoteEnumMapperTests
{
    [Theory]
    [InlineData("backup", OneNoteSpecialLocation.SlBackUpFolder)]
    [InlineData("unfiledNotes", OneNoteSpecialLocation.SlUnfiledNotesSection)]
    [InlineData("defaultNotebook", OneNoteSpecialLocation.SlDefaultNotebookFolder)]
    [InlineData("DEFAULTNOTEBOOK", OneNoteSpecialLocation.SlDefaultNotebookFolder)]
    public void MapSpecialLocation_MapsKnownNames(string name, int expected)
    {
        Assert.Equal(expected, OneNoteEnumMapper.MapSpecialLocation(name));
    }

    [Fact]
    public void MapSpecialLocation_Unknown_Throws()
    {
        Assert.Throws<ArgumentException>(() => OneNoteEnumMapper.MapSpecialLocation("nowhere"));
    }

    [Theory]
    [InlineData("email", OneNoteFilingLocation.FlEMail)]
    [InlineData("contacts", OneNoteFilingLocation.FlContacts)]
    [InlineData("tasks", OneNoteFilingLocation.FlTasks)]
    [InlineData("meetings", OneNoteFilingLocation.FlMeetings)]
    [InlineData("webContent", OneNoteFilingLocation.FlWebContent)]
    [InlineData("printOuts", OneNoteFilingLocation.FlPrintOuts)]
    public void MapFilingLocation_MapsKnownNames(string name, int expected)
    {
        Assert.Equal(expected, OneNoteEnumMapper.MapFilingLocation(name));
    }

    [Fact]
    public void MapFilingLocation_Unknown_Throws()
    {
        Assert.Throws<ArgumentException>(() => OneNoteEnumMapper.MapFilingLocation("faxes"));
    }

    [Theory]
    [InlineData("namedSectionNewPage", OneNoteFilingLocationType.FltNamedSectionNewPage)]
    [InlineData("currentSectionNewPage", OneNoteFilingLocationType.FltCurrentSectionNewPage)]
    [InlineData("currentPage", OneNoteFilingLocationType.FltCurrentPage)]
    // The interop skips 3 — namedPage is 4; guard that the mapping honors it.
    [InlineData("namedPage", OneNoteFilingLocationType.FltNamedPage)]
    public void MapFilingLocationType_MapsKnownNames(string name, int expected)
    {
        Assert.Equal(expected, OneNoteEnumMapper.MapFilingLocationType(name));
    }

    [Fact]
    public void MapFilingLocationType_NamedPage_Is4_NotContiguous()
    {
        Assert.Equal(4, OneNoteEnumMapper.MapFilingLocationType("namedPage"));
    }

    [Fact]
    public void MapFilingLocationType_Unknown_Throws()
    {
        Assert.Throws<ArgumentException>(() => OneNoteEnumMapper.MapFilingLocationType("someType"));
    }
}
