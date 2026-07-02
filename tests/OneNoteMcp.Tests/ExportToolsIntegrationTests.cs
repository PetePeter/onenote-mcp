using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// End-to-end COM integration tests for the PDF export tool. Each test publishes
/// under the shared fixture notebook into a scratch %TEMP% directory and cleans up
/// in a finally, so nothing leaks between runs. They early-return (skip) when the
/// fixture is unavailable. The COM-free argument-validation checks run everywhere.
/// </summary>
[Collection("OneNote COM")]
public sealed class ExportToolsIntegrationTests
{
    private readonly FixtureNotebook _fx;

    public ExportToolsIntegrationTests(FixtureNotebook fx) => _fx = fx;

    // ── single mode ────────────────────────────────────────────────────────────

    [Fact]
    public void ExportPdf_SectionSingleMode_ProducesPdf()
    {
        if (!_fx.Available) return;

        var target = ScratchFile("pdf");
        try
        {
            var json = ExportTools.ExportPdf(FixtureNotebook.TestVersion, _fx.SectionId, target, "single");

            var paths = Deserialize(json);
            Assert.Single(paths);
            Assert.Equal(target, paths[0]);
            AssertIsPdf(target);
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    [Fact]
    public void ExportPdf_SinglePage_ProducesPdf()
    {
        if (!_fx.Available) return;

        var target = ScratchFile("pdf");
        try
        {
            var json = ExportTools.ExportPdf(FixtureNotebook.TestVersion, _fx.TextPageId, target, "single");

            var paths = Deserialize(json);
            Assert.Single(paths);
            AssertIsPdf(paths[0]);
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    [Fact]
    public void ExportPdf_SingleMode_ExistingTargetIsOverwritten()
    {
        if (!_fx.Available) return;

        // Publish to an EXISTING path fails in OneNote; the tool must delete the
        // stale target first. Seed a bogus file and prove it becomes a real PDF.
        var target = ScratchFile("pdf");
        File.WriteAllText(target, "not a pdf");
        try
        {
            ExportTools.ExportPdf(FixtureNotebook.TestVersion, _fx.SectionId, target, "single");
            AssertIsPdf(target);
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    [Fact]
    public void ExportPdf_SectionRunTwiceSameSession_Succeeds()
    {
        if (!_fx.Available) return;

        // Acceptance criterion: repeated exports must not wedge OneNote.
        var target = ScratchFile("pdf");
        try
        {
            ExportTools.ExportPdf(FixtureNotebook.TestVersion, _fx.SectionId, target, "single");
            ExportTools.ExportPdf(FixtureNotebook.TestVersion, _fx.SectionId, target, "single");
            AssertIsPdf(target);
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    // ── perPage mode ───────────────────────────────────────────────────────────

    [Fact]
    public void ExportPdf_PerPageMode_ProducesOnePdfPerPage()
    {
        if (!_fx.Available) return;

        // The fixture section holds exactly two pages (text + image).
        var dir = ScratchDir();
        try
        {
            var json = ExportTools.ExportPdf(FixtureNotebook.TestVersion, _fx.SectionId, dir, "perPage");

            var paths = Deserialize(json);
            Assert.Equal(2, paths.Length);
            foreach (var p in paths) AssertIsPdf(p);
        }
        finally
        {
            TryDeleteDir(dir);
        }
    }

    [Fact]
    public void ExportPdf_PerPageMode_HonorsInterPublishDelay()
    {
        if (!_fx.Available) return;

        // Two pages ⇒ at least one inter-publish delay must elapse.
        const int delayMs = 1500;
        var dir = ScratchDir();
        try
        {
            var sw = Stopwatch.StartNew();
            ExportTools.ExportPdf(FixtureNotebook.TestVersion, _fx.SectionId, dir, "perPage", delayMs);
            sw.Stop();

            Assert.True(sw.ElapsedMilliseconds >= delayMs,
                $"Expected at least one {delayMs}ms delay; elapsed {sw.ElapsedMilliseconds}ms.");
        }
        finally
        {
            TryDeleteDir(dir);
        }
    }

    [Fact]
    public void ExportPdf_PerPageMode_EmptySection_ReturnsEmptyArray()
    {
        if (!_fx.Available) return;

        var sectionId = SectionTools.CreateSection(FixtureNotebook.TestVersion, _fx.NotebookId, "Empty " + Guid.NewGuid().ToString("N"));
        var dir = ScratchDir();
        try
        {
            var json = ExportTools.ExportPdf(FixtureNotebook.TestVersion, sectionId, dir, "perPage");
            Assert.Empty(Deserialize(json));
        }
        finally
        {
            TryDeleteDir(dir);
            try { SectionTools.DeleteNode(FixtureNotebook.TestVersion, sectionId); } catch { /* best-effort */ }
        }
    }

    [Fact]
    public void ExportPdf_SingleMode_InvalidTargetPath_ThrowsWithoutCrashing()
    {
        if (!_fx.Available) return;

        // An illegal path (embedded null character) must surface a clean exception,
        // not crash. This is deterministic across machines, unlike a drive letter
        // that may or may not be mapped on any given host.
        var illegalPath = Path.Combine(Path.GetTempPath(), "onenote-mcp\0nodir", "x.pdf");
        Assert.ThrowsAny<Exception>(
            () => ExportTools.ExportPdf(FixtureNotebook.TestVersion, _fx.SectionId, illegalPath, "single"));
    }

    // ── COM-free argument validation ─────────────────────────────────────────────

    [Fact]
    public void ExportPdf_UnknownMode_ThrowsArgumentExceptionBeforeAnySideEffect()
    {
        var target = ScratchFile("pdf");

        var ex = Assert.Throws<ArgumentException>(
            () => ExportTools.ExportPdf(FixtureNotebook.TestVersion, "{id}", target, "bogus"));

        Assert.Contains("bogus", ex.Message);
        Assert.False(File.Exists(target), "Validation must run before any file is produced.");
    }

    // ── helpers ──────────────────────────────────────────────────────────────────

    private static string[] Deserialize(string json) =>
        JsonSerializer.Deserialize<string[]>(json) ?? Array.Empty<string>();

    private static void AssertIsPdf(string path)
    {
        Assert.True(File.Exists(path), $"Expected PDF at {path}");
        var bytes = File.ReadAllBytes(path);
        Assert.True(bytes.Length > 0, "PDF is empty.");
        Assert.True(
            bytes.Length >= 4 && bytes[0] == (byte)'%' && bytes[1] == (byte)'P'
                && bytes[2] == (byte)'D' && bytes[3] == (byte)'F',
            "File does not start with the %PDF header.");
    }

    private static string ScratchFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"onenote-mcp-export-{Guid.NewGuid():N}.{ext}");

    private static string ScratchDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"onenote-mcp-export-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void TryDeleteFile(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* best-effort */ }
    }

    private static void TryDeleteDir(string dir)
    {
        try { if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true); } catch { /* best-effort */ }
    }
}
