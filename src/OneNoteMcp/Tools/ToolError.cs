using System;
using OneNoteMcp.Interop;

namespace OneNoteMcp.Tools;

/// <summary>
/// Guard that every COM-touching tool runs its body through so a failure becomes a
/// human-readable tool result (errors-in-results) instead of an unhandled exception
/// carrying a raw HRESULT.
/// </summary>
public static class ToolError
{
    /// <summary>Runs <paramref name="body"/>, returning its result or the mapped error text.</summary>
    public static string Guard(Func<string> body)
    {
        try
        {
            return body();
        }
        catch (Exception ex)
        {
            return ComErrorMapper.Describe(ex);
        }
    }
}
