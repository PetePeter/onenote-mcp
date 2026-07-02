using OneNoteMcp.Interop;

namespace OneNoteMcp.Model;

/// <summary>
/// Pure, COM-free resolution of a caller-friendly version token — a year ("2007"),
/// an Office major ("12"), or a raw CLSID — to the canonical CLSID string
/// ("B"-uppercase braces). There is NO default: unknown or empty tokens throw
/// (P-0545 DECISIONS — every tool passes an explicit version).
/// </summary>
public static class VersionResolver
{
    /// <summary>
    /// Resolves a version token to its canonical CLSID string. Known year/major
    /// tokens map via the catalog; any valid GUID (braced or bare, any case) is
    /// accepted and normalized to uppercase braces. Empty, whitespace, or
    /// unrecognised tokens throw <see cref="ArgumentException"/>.
    /// </summary>
    public static string Resolve(string token)
    {
        if (string.IsNullOrWhiteSpace(token))
            throw new ArgumentException("Version token must not be empty.", nameof(token));

        var trimmed = token.Trim();

        foreach (var known in OneNoteVersionCatalog.All)
        {
            if (known.Tokens.Any(t => string.Equals(t, trimmed, StringComparison.OrdinalIgnoreCase)))
                return known.Clsid;
        }

        // Accept any valid CLSID even when it isn't in the catalog, normalizing to
        // the canonical uppercase-braces form.
        if (Guid.TryParse(trimmed, out var guid))
            return guid.ToString("B").ToUpperInvariant();

        throw new ArgumentException(
            $"Unknown version token '{token}'. Expected a known year/major (e.g. 2007, 12, 2010, 14, 2016, 16) or a CLSID.",
            nameof(token));
    }
}
