namespace OneNoteMcp.Model;

/// <summary>
/// The 23 wrapped OneNote IApplication methods, one per method exposed by
/// OneNoteSession. The last six were added across IApplication2-4 (2010+) and are
/// absent on OneNote 2007 (v12).
/// </summary>
public enum Capability
{
    GetHierarchy,
    FindPages,
    GetPageContent,
    UpdatePageContent,
    OpenHierarchy,
    Publish,
    CreateNewPage,
    CloseNotebook,
    DeleteHierarchy,
    UpdateHierarchy,
    GetBinaryPageContent,
    DeletePageContent,
    GetHierarchyParent,
    GetSpecialLocation,
    NavigateTo,
    GetHyperlinkToObject,
    FindMeta,
    // --- IApplication2-4 additions (2010+), unsupported on v12 ---
    NavigateToUrl,
    GetWebHyperlinkToObject,
    MergeFiles,
    MergeSections,
    SyncHierarchy,
    SetFilingLocation,
}

/// <summary>
/// Pure, COM-free per-major capability table encoding the P-0545 spike result:
/// OneNote 2007 (v12) supports 17 of the 23 wrapped methods; the six
/// IApplication2-4 additions are absent. Modern majors (14/15/16) support all 23.
/// Unknown majors throw — there is no default.
/// </summary>
public static class OneNoteCapabilities
{
    // The six methods added in 2010+, absent on v12.
    private static readonly IReadOnlySet<Capability> ModernOnly = new HashSet<Capability>
    {
        Capability.NavigateToUrl,
        Capability.GetWebHyperlinkToObject,
        Capability.MergeFiles,
        Capability.MergeSections,
        Capability.SyncHierarchy,
        Capability.SetFilingLocation,
    };

    // v12 (2007) supports everything except the six modern-only additions.
    private static readonly IReadOnlySet<Capability> V12Supported =
        Enum.GetValues<Capability>().Where(c => !ModernOnly.Contains(c)).ToHashSet();

    /// <summary>
    /// Returns whether the given Office major supports the capability. Major 12
    /// consults the v12 set; 14/15/16 support all capabilities. Any other major
    /// throws <see cref="ArgumentException"/>.
    /// </summary>
    public static bool Supports(int major, Capability cap) => major switch
    {
        12 => V12Supported.Contains(cap),
        14 or 15 or 16 => true,
        _ => throw new ArgumentException($"Unknown OneNote major version '{major}'.", nameof(major)),
    };

    /// <summary>
    /// Returns the set of capabilities supported by the given Office major,
    /// consistent with <see cref="Supports"/>. Unknown majors throw.
    /// </summary>
    public static IReadOnlySet<Capability> SupportedBy(int major) => major switch
    {
        12 => new HashSet<Capability>(V12Supported),
        14 or 15 or 16 => new HashSet<Capability>(Enum.GetValues<Capability>()),
        _ => throw new ArgumentException($"Unknown OneNote major version '{major}'.", nameof(major)),
    };
}
