using System;
using System.IO;
using System.Text.Json;
using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// End-to-end COM integration tests for the OneNote-native package export tools:
/// onenote_export_one (a section as .one) and onenote_export_onepkg (a notebook as
/// .onepkg). Each test publishes under the shared fixture notebook into a scratch
/// %TEMP% path and cleans up in a finally, so nothing leaks between runs. They
/// early-return (skip) when the fixture is unavailable; the COM-free argument checks
/// run everywhere.
/// </summary>
[Collection("OneNote COM")]
public sealed class ExportPackageIntegrationTests
{
    private readonly FixtureNotebook _fx;

    public ExportPackageIntegrationTests(FixtureNotebook fx) => _fx = fx;

    // ── .one (section) ───────────────────────────────────────────────────────────

    [Fact]
    public void ExportOne_Section_ProducesNativeOneFile()
    {
        if (!_fx.Available) return;

        // Verify the produced file is a genuine OneNote 2010+ section file by its
        // format-GUID header. (A round-trip re-import is unreliable here: a freshly
        // opened section syncs its pages asynchronously, so its titles are not yet
        // present when the tool returns.)
        var target = ScratchFile("one");
        try
        {
            var json = ExportTools.ExportOne(FixtureNotebook.TestVersion, _fx.SectionId, target);

            var paths = Deserialize(json);
            Assert.Single(paths);
            Assert.Equal(target, paths[0]);
            AssertIsOneFile(target);
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    [Fact]
    public void ExportOne_ExistingTarget_IsOverwritten()
    {
        if (!_fx.Available) return;

        // Publish refuses to overwrite; the tool must delete the stale target first.
        var target = ScratchFile("one");
        File.WriteAllText(target, "not a .one file");
        try
        {
            ExportTools.ExportOne(FixtureNotebook.TestVersion, _fx.SectionId, target);
            AssertNonEmptyFile(target);
            // A real .one is far larger than our 15-byte placeholder.
            Assert.True(new FileInfo(target).Length > 100, "Expected a real .one file, not the placeholder.");
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    // ── .onepkg (notebook) ───────────────────────────────────────────────────────

    [Fact]
    public void ExportOnepkg_Notebook_ProducesCabPackage()
    {
        if (!_fx.Available) return;

        // A .onepkg is a Windows cabinet; assert the "MSCF" magic bytes.
        var target = ScratchFile("onepkg");
        try
        {
            var json = ExportTools.ExportOnepkg(FixtureNotebook.TestVersion, _fx.NotebookId, target);

            var paths = Deserialize(json);
            Assert.Single(paths);
            Assert.Equal(target, paths[0]);
            AssertIsCab(target);
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    [Fact]
    public void ExportOne_InvalidTargetPath_ThrowsWithoutCrashing()
    {
        if (!_fx.Available) return;

        // An embedded null makes an illegal path deterministically across machines.
        var illegalPath = Path.Combine(Path.GetTempPath(), "onenote-mcp\0nodir", "x.one");
        Assert.ThrowsAny<Exception>(() => ExportTools.ExportOne(FixtureNotebook.TestVersion, _fx.SectionId, illegalPath));
    }

    // ── v12 guard tests ─────────────────────────────────────────────────────────

    [Fact]
    public void ExportOne_Version2007_ThrowsNotSupportedWithClearMessage()
    {
        var target = ScratchFile("one");
        try
        {
            var ex = Assert.Throws<NotSupportedInVersionException>(
                () => ExportTools.ExportOne("2007", "{fake-section-id}", target));

            Assert.Equal(12, ex.VersionMajor);
            Assert.Equal("Publish", ex.MethodName);
            Assert.Contains("2007", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("2010", ex.Message);
            Assert.Contains("2010+", ex.Message);
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    [Fact]
    public void ExportOnepkg_Version2007_ThrowsNotSupportedWithClearMessage()
    {
        var target = ScratchFile("onepkg");
        try
        {
            var ex = Assert.Throws<NotSupportedInVersionException>(
                () => ExportTools.ExportOnepkg("2007", "{fake-notebook-id}", target));

            Assert.Equal(12, ex.VersionMajor);
            Assert.Equal("Publish", ex.MethodName);
            Assert.Contains("2007", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("2010", ex.Message);
            Assert.Contains("2010+", ex.Message);
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    [Fact]
    public void ExportOne_VersionMajor12_ThrowsNotSupportedWithClearMessage()
    {
        // Verify major 12 (the numeric form of 2007) is also guarded.
        var target = ScratchFile("one");
        try
        {
            var ex = Assert.Throws<NotSupportedInVersionException>(
                () => ExportTools.ExportOne("12", "{fake-section-id}", target));

            Assert.Equal(12, ex.VersionMajor);
        }
        finally
        {
            TryDeleteFile(target);
        }
    }

    // ── helpers ──────────────────────────────────────────────────────────────────

    private static string[] Deserialize(string json) =>
        JsonSerializer.Deserialize<string[]>(json) ?? Array.Empty<string>();

    private static void AssertNonEmptyFile(string path)
    {
        Assert.True(File.Exists(path), $"Expected a file at {path}");
        Assert.True(new FileInfo(path).Length > 0, "File is empty.");
    }

    // The 16-byte guidFileType that opens every OneNote 2010+ .one section file:
    // {7B5C52E4-D8C8-4DA7-AEB1-5378D02996D3}, laid out in file byte order.
    private static readonly byte[] OneFileMagic =
    {
        0xE4, 0x52, 0x5C, 0x7B, 0x8C, 0xD8, 0xA7, 0x4D,
        0xAE, 0xB1, 0x53, 0x78, 0xD0, 0x29, 0x96, 0xD3,
    };

    private static void AssertIsOneFile(string path)
    {
        Assert.True(File.Exists(path), $"Expected a .one file at {path}");
        var bytes = File.ReadAllBytes(path);
        Assert.True(bytes.Length >= OneFileMagic.Length, "File is shorter than the .one header.");
        Assert.True(
            bytes.AsSpan(0, OneFileMagic.Length).SequenceEqual(OneFileMagic),
            "File does not start with the OneNote .one format-GUID header.");
    }

    private static void AssertIsCab(string path)
    {
        Assert.True(File.Exists(path), $"Expected a package at {path}");
        var bytes = File.ReadAllBytes(path);
        Assert.True(
            bytes.Length >= 4 && bytes[0] == (byte)'M' && bytes[1] == (byte)'S'
                && bytes[2] == (byte)'C' && bytes[3] == (byte)'F',
            "File does not start with the MSCF cabinet header.");
    }

    private static string ScratchFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"onenote-mcp-export-{Guid.NewGuid():N}.{ext}");

    private static void TryDeleteFile(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* best-effort */ }
    }
}
