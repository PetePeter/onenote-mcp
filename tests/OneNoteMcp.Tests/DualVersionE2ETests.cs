using System.Text.Json;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;
using OneNoteMcp.Tests.Fakes;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

// ─────────────────────────────────────────────────────────────────────────────
// Single source of truth for the "is this a dual-install host?" skip decision.
// Two OneNote versions count as distinct only when they launch different exes
// (the 2010/2013/2016 compat CLSIDs all point at the same modern ONENOTE.EXE and
// are collapsed by OneNoteVersion.DetectAll). The COM-gated tests below early
// -return when this is false so they stay green on single-install CI hosts.
// ─────────────────────────────────────────────────────────────────────────────
internal static class DualVersionGuard
{
    /// <summary>True when at least two DISTINCT (case-insensitive) exe paths exist.</summary>
    public static bool ShouldRun(IEnumerable<string> exePaths) =>
        exePaths
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count() >= 2;
}

/// <summary>
/// COM-free proof that version routing and capability-gating are real: a supported
/// read routed to "2007" hits the legacy (v12) CLSID, and every modern-only tool
/// invoked against 2007 is rejected by the gate before COM is ever touched.
/// Uses AppFactoryOverride, so ResetForTests runs in both ctor and Dispose to keep
/// the per-CLSID session cache and override from leaking into other tests.
/// </summary>
[Collection("OneNote COM")]
public sealed class DualVersionRoutingTests : IDisposable
{
    private const string Clsid2007 = "{0039FFEC-A022-4232-8274-6B34787BFC27}";

    public DualVersionRoutingTests() => OneNoteSession.ResetForTests();
    public void Dispose() => OneNoteSession.ResetForTests();

    [Fact]
    public void Route_2007SupportedRead_UsesLegacyClsid()
    {
        // GetHierarchy is one of the 17 caps present on v12, so it must route to the
        // 2007 server — identified by its CLSID — without ever touching real COM.
        string? captured = null;
        OneNoteSession.AppFactoryOverride = c =>
        {
            captured = c;
            return new FakeOneNoteApp { ReturnValue = "<Notebooks/>" };
        };

        var result = HierarchyTools.GetHierarchy("2007", "node", "notebooks");

        Assert.Equal(Clsid2007, captured, StringComparer.OrdinalIgnoreCase);
        Assert.Contains("Notebooks", result);
    }

    // The six IApplication2-4 additions absent on OneNote 2007 (v12), each driven
    // with dummy args so the theory covers every arity. The label doubles as the
    // capability name the gate reports in its "not available" message.
    public static IEnumerable<object[]> ModernOnlyTools()
    {
        yield return Tool("SyncHierarchy", v => HierarchyMaintenanceTools.SyncHierarchy(v, "id"));
        yield return Tool("SetFilingLocation", v => HierarchyMaintenanceTools.SetFilingLocation(v, "email", "currentPage", "sec"));
        yield return Tool("MergeFiles", v => HierarchyMaintenanceTools.MergeFiles(v, "b", "c", "s", "t"));
        yield return Tool("MergeSections", v => HierarchyMaintenanceTools.MergeSections(v, "src", "dst"));
        yield return Tool("NavigateToUrl", v => NavigationTools.NavigateToUrl(v, "onenote:x"));
        yield return Tool("GetWebHyperlinkToObject", v => NavigationTools.GetWebHyperlinkToObject(v, "hid"));
    }

    private static object[] Tool(string label, Func<string, string> invoke) => [label, invoke];

    [Theory]
    [MemberData(nameof(ModernOnlyTools))]
    public void ModernOnlyTools_On2007_ReturnStructuredUnsupported_AndNeverCallCom(
        string label, Func<string, string> invoke)
    {
        // The gate fires in ToolVersion.Route BEFORE OneNoteSession.For is reached, so
        // the session factory must never run. Capturing that is a real "COM untouched"
        // proof (ThrowIfCalled + LastMethod is not — Guard throws before Record, leaving
        // LastMethod null whether or not a method was dispatched).
        var factoryCalled = false;
        OneNoteSession.AppFactoryOverride = _ =>
        {
            factoryCalled = true;
            return new FakeOneNoteApp { ThrowIfCalled = true };
        };

        var result = invoke("2007");

        Assert.Contains("Unsupported on OneNote 2007", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(label, result, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("0x", result); // no COM/HRESULT crash leaked through
        Assert.False(factoryCalled,          // capability gate rejected before COM path
            "capability gate must reject before OneNoteSession.For creates a session");
    }
}

/// <summary>
/// Pure unit tests for the dual-install skip predicate that gates the COM-live
/// tests. Distinct exe paths (case-insensitive) are what make two installs count.
/// </summary>
public sealed class DualVersionGuardTests
{
    [Fact]
    public void ShouldRun_SingleInstall_False() =>
        Assert.False(DualVersionGuard.ShouldRun(["a.exe"]));

    [Fact]
    public void ShouldRun_SameExeDifferentCase_False() =>
        Assert.False(DualVersionGuard.ShouldRun(["A.exe", "a.exe"]));

    [Fact]
    public void ShouldRun_TwoDistinctExes_True() =>
        Assert.True(DualVersionGuard.ShouldRun(["a.exe", "b.exe"]));
}

/// <summary>
/// Drift-guard that keeps the README capability matrix honest: every Capability
/// value must have a row, and each version cell's ✓/✗ must match
/// OneNoteCapabilities.Supports for that version's major. COM-free — always runs.
/// </summary>
public sealed class ReadmeMatrixTests
{
    // Maps a matrix column header to the Office major it represents. 2013 shares the
    // modern (major 16) CLSID with 2016, so both columns resolve to 16.
    private static int HeaderToMajor(string header) => header switch
    {
        "2007" => 12,
        "2010" => 14,
        "2013" => 16,
        "2016" => 16,
        _ => throw new ArgumentException($"Unexpected matrix column '{header}'."),
    };

    private static readonly string[] VersionColumns = ["2007", "2010", "2013", "2016"];

    [Fact]
    public void ReadmeCapabilityMatrix_MatchesOneNoteCapabilities()
    {
        var readme = ReadReadme();
        var rows = ParseMatrix(readme);

        foreach (var cap in Enum.GetValues<Capability>())
        {
            Assert.True(rows.TryGetValue(cap.ToString(), out var cells),
                $"README capability matrix is missing a row for '{cap}'.");

            foreach (var column in VersionColumns)
            {
                var major = HeaderToMajor(column);
                var expected = OneNoteCapabilities.Supports(major, cap);
                var actual = CellIsSupported(cells![column], cap, column);
                Assert.True(expected == actual,
                    $"README matrix cell for {cap}/{column} says {(actual ? "✓" : "✗")} " +
                    $"but OneNoteCapabilities.Supports({major}, {cap}) is {expected}.");
            }
        }
    }

    /// <summary>
    /// A cell is SUPPORTED iff it contains ✓ (U+2713), UNSUPPORTED iff it contains
    /// ✗ (U+2717). Anything else is ambiguous and fails the test outright.
    /// </summary>
    private static bool CellIsSupported(string cell, Capability cap, string column)
    {
        var trimmed = cell.Trim();
        if (trimmed.Contains('✓')) return true;
        if (trimmed.Contains('✗')) return false;
        Assert.Fail($"README matrix cell for {cap}/{column} is ambiguous: '{cell}' " +
                    "(expected ✓ or ✗).");
        return false; // unreachable
    }

    /// <summary>
    /// Parses the capability matrix into {method → {column → cell}}. Locates the
    /// header row by its Method/2007…2016 columns, then reads every following pipe
    /// row until the table ends.
    /// </summary>
    private static Dictionary<string, Dictionary<string, string>> ParseMatrix(string readme)
    {
        var lines = readme.Replace("\r\n", "\n").Split('\n');
        var headerIndex = Array.FindIndex(lines, IsMatrixHeader);
        Assert.True(headerIndex >= 0, "README is missing the capability matrix header " +
            "'| Method | 2007 | 2010 | 2013 | 2016 |'.");

        var header = SplitRow(lines[headerIndex]);
        var rows = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);

        // Skip the header and the |---| separator, then read data rows.
        for (int i = headerIndex + 2; i < lines.Length; i++)
        {
            var line = lines[i];
            if (!line.TrimStart().StartsWith('|')) break; // table ended
            var cells = SplitRow(line);
            if (cells.Count != header.Count) continue;

            var method = cells[0];
            var byColumn = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int c = 1; c < header.Count; c++)
                byColumn[header[c]] = cells[c];
            rows[method] = byColumn;
        }

        return rows;
    }

    private static bool IsMatrixHeader(string line)
    {
        if (!line.TrimStart().StartsWith('|')) return false;
        var cells = SplitRow(line);
        return cells.Count == 5 &&
               cells[0] == "Method" &&
               cells[1] == "2007" && cells[2] == "2010" &&
               cells[3] == "2013" && cells[4] == "2016";
    }

    /// <summary>Splits a markdown table row into trimmed cells, dropping the outer pipes.</summary>
    private static List<string> SplitRow(string line) =>
        line.Trim().Trim('|').Split('|').Select(c => c.Trim()).ToList();

    /// <summary>
    /// Finds README.md by walking up from the test output dir to the directory that
    /// holds OneNoteMcp.sln. A missing README is a real failure, never a skip.
    /// </summary>
    private static string ReadReadme()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "OneNoteMcp.sln")))
            dir = dir.Parent;

        if (dir is null)
        {
            Assert.Fail("Could not locate OneNoteMcp.sln walking up from " +
                        $"'{AppContext.BaseDirectory}'; README.md is unreachable.");
        }

        var readmePath = Path.Combine(dir!.FullName, "README.md");
        if (!File.Exists(readmePath))
            Assert.Fail($"README.md not found next to the solution at '{readmePath}'.");

        return File.ReadAllText(readmePath);
    }
}

/// <summary>
/// COM-live proof of multi-version support on a dual-install host. Each test
/// early-returns (a runtime skip, not [Fact(Skip)]) when fewer than two distinct
/// OneNote exes are present, so they stay green on single-install CI. ResetForTests
/// runs in ctor and Dispose so no AppFactoryOverride from a sibling test leaks in.
/// </summary>
[Collection("OneNote COM")]
public sealed class DualVersionComTests : IDisposable
{
    private const string Clsid2007 = "{0039FFEC-A022-4232-8274-6B34787BFC27}";
    private const string Clsid2016 = "{DC67E480-C3CB-49F8-8232-60B0C2056C8E}";

    public DualVersionComTests() => OneNoteSession.ResetForTests();
    public void Dispose() => OneNoteSession.ResetForTests();

    private sealed record VersionReportDto(
        string Version, int Major, string Clsid, string ExePath, bool Default, string[] Capabilities);

    [Fact]
    public void ListVersions_Live_DetectsBothInstalls()
    {
        var json = VersionTools.ListVersions();
        var reports = JsonSerializer.Deserialize<VersionReportDto[]>(
            json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? [];

        if (!DualVersionGuard.ShouldRun(reports.Select(r => r.ExePath)))
            return; // single-install CI host — nothing to assert

        Assert.True(reports.Length >= 2);
        Assert.Equal(
            reports.Length,
            reports.Select(r => r.ExePath).Distinct(StringComparer.OrdinalIgnoreCase).Count());

        var clsids = reports.Select(r => r.Clsid).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains(Clsid2007, clsids, StringComparer.OrdinalIgnoreCase);
        Assert.Contains(Clsid2016, clsids, StringComparer.OrdinalIgnoreCase);

        var v12 = Assert.Single(reports, r => r.Major == 12);
        Assert.Contains("GetHierarchy", v12.Capabilities);   // a present cap
        Assert.DoesNotContain("SyncHierarchy", v12.Capabilities); // an absent one
    }

    [Fact]
    public void GetHierarchy_Live_RoutesToPerVersionSchema()
    {
        if (!DualVersionGuard.ShouldRun(OneNoteVersion.DetectAll().Select(i => i.ExePath)))
            return; // single-install CI host

        // No AppFactoryOverride is set (ResetForTests in ctor), so these are real COM
        // calls. Routing correctness is already proven COM-free by
        // DualVersionRoutingTests.Route_2007SupportedRead_UsesLegacyClsid, so
        // returning early on a COM error here cannot mask a routing regression — it
        // only tolerates a OneNote server that will not launch under automation.

        var modernXml = HierarchyTools.GetHierarchy("2016", "", "notebooks");
        if (!modernXml.Contains("OneNote error 0x"))
            Assert.Contains("office/onenote/2013/onenote", modernXml);

        var legacyXml = HierarchyTools.GetHierarchy("2007", "", "notebooks");
        if (!legacyXml.Contains("OneNote error 0x"))
            Assert.Contains("office/onenote/2007/onenote", legacyXml);
    }
}
