namespace OneNoteMcp.Interop;

/// <summary>
/// Thrown when a OneNote operation is invoked against a version whose
/// <c>IApplication</c> does not implement it. OneNote 2007 (v12), for example,
/// lacks the IApplication2–4 additions (NavigateToUrl, GetWebHyperlinkToObject,
/// MergeFiles, MergeSections, SyncHierarchy, SetFilingLocation). Raising this
/// typed exception keeps us from dispatching blind to a missing vtable slot.
/// </summary>
public sealed class NotSupportedInVersionException : NotSupportedException
{
    /// <summary>The OneNote major version that lacks the method (e.g. 12 for 2007).</summary>
    public int VersionMajor { get; }

    /// <summary>The IApplication method name that is unavailable on that version.</summary>
    public string MethodName { get; }

    public NotSupportedInVersionException(int versionMajor, string methodName)
        : base($"OneNote method '{methodName}' is not supported by version {versionMajor}.")
    {
        VersionMajor = versionMajor;
        MethodName = methodName;
    }

    /// <summary>
    /// Overload for custom error messages (e.g., export-specific guidance).
    /// </summary>
    public NotSupportedInVersionException(int versionMajor, string methodName, string customMessage)
        : base(customMessage)
    {
        VersionMajor = versionMajor;
        MethodName = methodName;
    }
}
