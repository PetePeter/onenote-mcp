using OneNoteMcp.Interop;

namespace OneNoteMcp.Model;

/// <summary>
/// Pure, COM-free mapping of caller-friendly names to OneNote enum integer values
/// for the location / filing tools, mirroring <see cref="PageInfoMapper.MapDetail"/>.
/// Case-insensitive; an unknown name throws <see cref="ArgumentException"/> listing
/// the accepted values.
/// </summary>
public static class OneNoteEnumMapper
{
    /// <summary>Maps a friendly special-location name to its SpecialLocation value.</summary>
    public static int MapSpecialLocation(string location) => (location ?? "").ToLowerInvariant() switch
    {
        "backup" or "backupfolder"               => OneNoteSpecialLocation.SlBackUpFolder,
        "unfilednotes" or "unfilednotessection"  => OneNoteSpecialLocation.SlUnfiledNotesSection,
        "defaultnotebook" or "defaultnotebookfolder" => OneNoteSpecialLocation.SlDefaultNotebookFolder,
        _ => throw new ArgumentException(
            $"Unknown location '{location}'. Expected one of: backup, unfiledNotes, defaultNotebook.",
            nameof(location)),
    };

    /// <summary>Maps a friendly filing-location name to its FilingLocation value.</summary>
    public static int MapFilingLocation(string location) => (location ?? "").ToLowerInvariant() switch
    {
        "email"      => OneNoteFilingLocation.FlEMail,
        "contacts"   => OneNoteFilingLocation.FlContacts,
        "tasks"      => OneNoteFilingLocation.FlTasks,
        "meetings"   => OneNoteFilingLocation.FlMeetings,
        "webcontent" => OneNoteFilingLocation.FlWebContent,
        "printouts"  => OneNoteFilingLocation.FlPrintOuts,
        _ => throw new ArgumentException(
            $"Unknown filing location '{location}'. Expected one of: email, contacts, tasks, meetings, webContent, printOuts.",
            nameof(location)),
    };

    /// <summary>Maps a friendly filing-location-type name to its FilingLocationType value.</summary>
    public static int MapFilingLocationType(string locationType) => (locationType ?? "").ToLowerInvariant() switch
    {
        "namedsectionnewpage"   => OneNoteFilingLocationType.FltNamedSectionNewPage,
        "currentsectionnewpage" => OneNoteFilingLocationType.FltCurrentSectionNewPage,
        "currentpage"           => OneNoteFilingLocationType.FltCurrentPage,
        "namedpage"             => OneNoteFilingLocationType.FltNamedPage,
        _ => throw new ArgumentException(
            $"Unknown filing location type '{locationType}'. Expected one of: namedSectionNewPage, currentSectionNewPage, currentPage, namedPage.",
            nameof(locationType)),
    };
}
