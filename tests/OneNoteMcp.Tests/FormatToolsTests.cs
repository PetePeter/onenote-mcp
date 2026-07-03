using System;
using System.IO;
using System.Text.Json;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// COM-free tests for onenote_detect_format. Detection is a pure file-header read,
/// so these run everywhere against clean-room synthetic .one headers written to
/// %TEMP% (no real content, no COM). They also pin that detection never modifies
/// the file it inspects.
/// </summary>
public class FormatToolsDetectTests
{
    private static readonly Guid LegacyFormatGuid = new("11111111-2222-3333-4444-555555555555");

    [Fact]
    public void DetectFormat_CurrentSectionFile_ReportsCurrent()
    {
        var path = WriteHeaderFile(
            OneFileFormatSniffer.FileTypeSection,
            OneFileFormatSniffer.FileFormatRevisionStore);
        try
        {
            var json = FormatTools.DetectFormat(FixtureNotebook.TestVersion, path);
            var doc = Parse(json);

            Assert.Equal("current2010Plus", doc.GetProperty("format").GetString());
            Assert.Equal("section", doc.GetProperty("fileType").GetString());
            Assert.True(doc.GetProperty("isCurrent").GetBoolean());
            Assert.False(doc.GetProperty("is2007").GetBoolean());
        }
        finally
        {
            TryDelete(path);
        }
    }

    [Fact]
    public void DetectFormat_LegacySectionFile_ReportsIs2007()
    {
        var path = WriteHeaderFile(OneFileFormatSniffer.FileTypeSection, LegacyFormatGuid);
        try
        {
            var json = FormatTools.DetectFormat(FixtureNotebook.TestVersion, path);
            var doc = Parse(json);

            Assert.Equal("legacy2007", doc.GetProperty("format").GetString());
            Assert.True(doc.GetProperty("is2007").GetBoolean());
            Assert.False(doc.GetProperty("isCurrent").GetBoolean());
        }
        finally
        {
            TryDelete(path);
        }
    }

    [Fact]
    public void DetectFormat_NeverModifiesTheFile()
    {
        var path = WriteHeaderFile(
            OneFileFormatSniffer.FileTypeSection,
            OneFileFormatSniffer.FileFormatRevisionStore);
        var before = File.ReadAllBytes(path);
        try
        {
            FormatTools.DetectFormat(FixtureNotebook.TestVersion, path);

            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            TryDelete(path);
        }
    }

    [Fact]
    public void DetectFormat_MissingPath_ReturnsReportWithoutCrashing()
    {
        // A path that is neither an existing file nor a resolvable OneNote node must
        // produce an honest report, never an unhandled crash.
        var missing = Path.Combine(Path.GetTempPath(), $"detect-missing-{Guid.NewGuid():N}.one");
        var json = FormatTools.DetectFormat(FixtureNotebook.TestVersion, missing);
        var doc = Parse(json);
        Assert.Equal("notOneNoteFile", doc.GetProperty("format").GetString());
        Assert.Equal(missing, doc.GetProperty("input").GetString());
        Assert.False(string.IsNullOrWhiteSpace(doc.GetProperty("detail").GetString()));
    }

    private static string WriteHeaderFile(Guid fileType, Guid fileFormat)
    {
        var header = new byte[OneFileFormatSniffer.HeaderLength];
        fileType.ToByteArray().CopyTo(header, 0);
        fileFormat.ToByteArray().CopyTo(header, 48);
        var path = Path.Combine(Path.GetTempPath(), $"detect-{Guid.NewGuid():N}.one");
        File.WriteAllBytes(path, header);
        return path;
    }

    private static JsonElement Parse(string json) => JsonDocument.Parse(json).RootElement;

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* best-effort */ }
    }
}

/// <summary>
/// COM integration tests for onenote_convert_section. Converting the fixture's
/// current-format section is a no-op-style success that proves the tool republishes
/// a genuine current-format .one; a bogus id proves the tool reports failure per
/// section instead of crashing. These skip when the fixture is unavailable. We never
/// touch any real notebook.
/// </summary>
[Collection("OneNote COM")]
public sealed class FormatToolsConvertIntegrationTests
{
    private readonly Fixtures.FixtureNotebook _fx;

    public FormatToolsConvertIntegrationTests(Fixtures.FixtureNotebook fx) => _fx = fx;

    [Fact]
    public void ConvertSection_FixtureSection_ProducesCurrentFormatOneFile()
    {
        if (!_fx.Available) return;

        var target = ScratchFile("one");
        try
        {
            var json = FormatTools.ConvertSection(FixtureNotebook.TestVersion, _fx.SectionId, target);
            var doc = JsonDocument.Parse(json).RootElement;

            Assert.True(doc.GetProperty("success").GetBoolean());
            Assert.True(File.Exists(target), "Converted .one file should exist.");
            // The produced file must itself sniff as current 2010+ format.
            Assert.Equal(OneNoteFileFormat.Current2010Plus, OneFileFormatSniffer.SniffFile(target).Format);
        }
        finally
        {
            TryDelete(target);
        }
    }

    [Fact]
    public void ConvertSection_BogusId_ReportsFailureWithoutCrashing()
    {
        if (!_fx.Available) return;

        var target = ScratchFile("one");
        try
        {
            var json = FormatTools.ConvertSection(FixtureNotebook.TestVersion, "{00000000-0000-0000-0000-000000000000}{1}{0}", target);
            var doc = JsonDocument.Parse(json).RootElement;

            Assert.False(doc.GetProperty("success").GetBoolean());
            Assert.False(string.IsNullOrWhiteSpace(doc.GetProperty("reason").GetString()));
        }
        finally
        {
            TryDelete(target);
        }
    }

    [Fact]
    public void ConvertSection_Version2007_ReportsFailureWithClearMessage()
    {
        // Section conversion from 2007 should fail early with a clear message,
        // before attempting any COM operations. ConvertSection never throws; it
        // always returns a JSON report for batch use-cases.
        var target = ScratchFile("one");
        try
        {
            var json = FormatTools.ConvertSection("2007", "{fake-section-id}", target);
            var doc = JsonDocument.Parse(json).RootElement;

            Assert.False(doc.GetProperty("success").GetBoolean());
            var reason = doc.GetProperty("reason").GetString() ?? "";
            Assert.Contains("2007", reason, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("2010", reason);
            Assert.Contains("conversion", reason, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            TryDelete(target);
        }
    }

    [Fact]
    public void ConvertSection_VersionMajor12_ReportsFailureWithClearMessage()
    {
        // Verify major 12 (the numeric form of 2007) is also guarded.
        var target = ScratchFile("one");
        try
        {
            var json = FormatTools.ConvertSection("12", "{fake-section-id}", target);
            var doc = JsonDocument.Parse(json).RootElement;

            Assert.False(doc.GetProperty("success").GetBoolean());
            Assert.Contains("2007", doc.GetProperty("reason").GetString() ?? "", StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            TryDelete(target);
        }
    }

    private static string ScratchFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"onenote-mcp-convert-{Guid.NewGuid():N}.{ext}");

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* best-effort */ }
    }
}
