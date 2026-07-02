using System;
using System.IO;
using OneNoteMcp.Model;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Pure, COM-free unit tests for the OneNote file-format sniffer. They assert the
/// classification of the 64-byte MS-ONESTORE header purely from bytes, using
/// clean-room synthetic headers (correct guidFileType / guidFileFormat, NO real
/// content) built in-repo — nothing is read from any real notebook.
///
/// Key spec fact (MS-ONESTORE section 2.3.1): the version discriminator is
/// guidFileFormat at offset 48, which is {109ADD3F-911B-49F5-A5D0-1791EDC8AED8}
/// for the documented 2010+ revision-store format. guidFileType at offset 0 only
/// says section (.one) vs table-of-contents (.onetoc2), NOT the version.
/// </summary>
public class OneFileFormatSnifferTests
{
    // A synthetic, clean-room guidFileFormat that is NOT the 2010+ revision store —
    // stands in for a legacy/2007 file whose format GUID differs. (Undocumented by
    // Microsoft, so any non-revision-store format GUID is classified legacy.)
    private static readonly Guid LegacyFormatGuid = new("11111111-2222-3333-4444-555555555555");

    // ── current 2010+ ────────────────────────────────────────────────────────────

    [Fact]
    public void Sniff_CurrentSection_ClassifiesCurrent2010Plus()
    {
        var header = BuildHeader(
            OneFileFormatSniffer.FileTypeSection,
            OneFileFormatSniffer.FileFormatRevisionStore);

        var result = OneFileFormatSniffer.Sniff(header);

        Assert.Equal(OneNoteFileFormat.Current2010Plus, result.Format);
        Assert.Equal("section", result.FileType);
        Assert.True(result.IsCurrent);
        Assert.False(result.Is2007);
        Assert.True(result.IsOneNoteFile);
    }

    [Fact]
    public void Sniff_CurrentTableOfContents_ClassifiesCurrentAndToc()
    {
        var header = BuildHeader(
            OneFileFormatSniffer.FileTypeTableOfContents,
            OneFileFormatSniffer.FileFormatRevisionStore);

        var result = OneFileFormatSniffer.Sniff(header);

        Assert.Equal(OneNoteFileFormat.Current2010Plus, result.Format);
        Assert.Equal("tableOfContents", result.FileType);
        Assert.True(result.IsCurrent);
    }

    // ── legacy / 2007 ──────────────────────────────────────────────────────────────

    [Fact]
    public void Sniff_LegacySection_ClassifiesLegacy2007()
    {
        // Valid OneNote section guidFileType, but a non-revision-store guidFileFormat.
        var header = BuildHeader(OneFileFormatSniffer.FileTypeSection, LegacyFormatGuid);

        var result = OneFileFormatSniffer.Sniff(header);

        Assert.Equal(OneNoteFileFormat.Legacy2007, result.Format);
        Assert.Equal("section", result.FileType);
        Assert.True(result.Is2007);
        Assert.False(result.IsCurrent);
        Assert.True(result.IsOneNoteFile);
    }

    // ── not a OneNote file ─────────────────────────────────────────────────────────

    [Fact]
    public void Sniff_UnknownFileType_ClassifiesNotOneNoteFile()
    {
        // A random guidFileType that matches neither .one nor .onetoc2.
        var header = BuildHeader(Guid.NewGuid(), OneFileFormatSniffer.FileFormatRevisionStore);

        var result = OneFileFormatSniffer.Sniff(header);

        Assert.Equal(OneNoteFileFormat.NotOneNoteFile, result.Format);
        Assert.Equal("unknown", result.FileType);
        Assert.False(result.IsOneNoteFile);
    }

    [Fact]
    public void Sniff_RawOnDiskSpecBytes_ClassifiesCurrentSection()
    {
        // Endianness guard: hand-written on-disk bytes straight from MS-ONESTORE,
        // independent of Guid.ToByteArray(), so a byte-order bug in the sniffer
        // can't hide behind a symmetric round-trip. guidFileType .one at offset 0,
        // guidFileFormat revision-store {109ADD3F-911B-49F5-A5D0-1791EDC8AED8} at 48.
        var header = new byte[64];
        // {7B5C52E4-D88C-4DA7-AEB1-5378D02996D3}
        byte[] section =
        {
            0xE4, 0x52, 0x5C, 0x7B, 0x8C, 0xD8, 0xA7, 0x4D,
            0xAE, 0xB1, 0x53, 0x78, 0xD0, 0x29, 0x96, 0xD3,
        };
        // {109ADD3F-911B-49F5-A5D0-1791EDC8AED8}
        byte[] revisionStore =
        {
            0x3F, 0xDD, 0x9A, 0x10, 0x1B, 0x91, 0xF5, 0x49,
            0xA5, 0xD0, 0x17, 0x91, 0xED, 0xC8, 0xAE, 0xD8,
        };
        section.CopyTo(header, 0);
        revisionStore.CopyTo(header, 48);

        var result = OneFileFormatSniffer.Sniff(header);

        Assert.Equal(OneNoteFileFormat.Current2010Plus, result.Format);
        Assert.Equal("section", result.FileType);
    }

    [Fact]
    public void Sniff_HeaderTooShort_ClassifiesNotOneNoteFile()
    {
        // Only the guidFileType is present; the guidFileFormat at offset 48 is missing.
        var tooShort = new byte[16];
        OneFileFormatSniffer.FileTypeSection.ToByteArray().CopyTo(tooShort, 0);

        var result = OneFileFormatSniffer.Sniff(tooShort);

        Assert.Equal(OneNoteFileFormat.NotOneNoteFile, result.Format);
        Assert.False(result.IsOneNoteFile);
    }

    // ── file-level read (never modifies) ────────────────────────────────────────────

    [Fact]
    public void SniffFile_ReadsCurrentSection_AndDoesNotModifyTheFile()
    {
        var path = Path.Combine(Path.GetTempPath(), $"sniff-{Guid.NewGuid():N}.one");
        var bytes = BuildHeader(
            OneFileFormatSniffer.FileTypeSection,
            OneFileFormatSniffer.FileFormatRevisionStore);
        File.WriteAllBytes(path, bytes);
        try
        {
            var result = OneFileFormatSniffer.SniffFile(path);

            Assert.Equal(OneNoteFileFormat.Current2010Plus, result.Format);
            // Read-only: bytes are untouched (authoritative check; no timestamp
            // assertion — coarse FAT/exFAT mtime granularity makes it flaky).
            Assert.Equal(bytes, File.ReadAllBytes(path));
        }
        finally
        {
            TryDelete(path);
        }
    }

    [Fact]
    public void SniffFile_MissingFile_Throws()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"sniff-missing-{Guid.NewGuid():N}.one");
        Assert.Throws<FileNotFoundException>(() => OneFileFormatSniffer.SniffFile(missing));
    }

    // ── helpers ──────────────────────────────────────────────────────────────────

    /// <summary>
    /// Builds a clean-room 64-byte MS-ONESTORE header: guidFileType at offset 0 and
    /// guidFileFormat at offset 48. Guid.ToByteArray() yields the exact on-disk
    /// mixed-endian byte layout, so Sniff can round-trip it via new Guid(span).
    /// </summary>
    private static byte[] BuildHeader(Guid fileType, Guid fileFormat)
    {
        var header = new byte[OneFileFormatSniffer.HeaderLength];
        fileType.ToByteArray().CopyTo(header, 0);
        fileFormat.ToByteArray().CopyTo(header, 48);
        return header;
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* best-effort */ }
    }
}
