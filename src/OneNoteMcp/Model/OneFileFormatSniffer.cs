using System;
using System.IO;

namespace OneNoteMcp.Model;

/// <summary>Version classification of a OneNote store file, derived from its header.</summary>
public enum OneNoteFileFormat
{
    /// <summary>The bytes are not a recognisable OneNote store file.</summary>
    NotOneNoteFile,

    /// <summary>A OneNote file that is not the documented 2010+ revision-store format.</summary>
    Legacy2007,

    /// <summary>The documented 2010+ revision-store format.</summary>
    Current2010Plus,
}

/// <summary>
/// Immutable outcome of sniffing a OneNote header: the version classification, the
/// file-type label ("section"/"tableOfContents"/"unknown") and a human-readable
/// detail explaining WHY the classification was chosen (useful in reports/logs).
/// </summary>
public sealed record OneFileFormatResult(OneNoteFileFormat Format, string FileType, string Detail)
{
    /// <summary>True for a legacy/2007-era OneNote file.</summary>
    public bool Is2007 => Format == OneNoteFileFormat.Legacy2007;

    /// <summary>True for the current 2010+ revision-store format.</summary>
    public bool IsCurrent => Format == OneNoteFileFormat.Current2010Plus;

    /// <summary>True when the bytes are a recognisable OneNote store file at all.</summary>
    public bool IsOneNoteFile => Format != OneNoteFileFormat.NotOneNoteFile;
}

/// <summary>
/// Pure, COM-free classifier for the 64-byte MS-ONESTORE file header (spec section
/// 2.3.1). It reads ONLY two header GUIDs — guidFileType at offset 0 (section vs
/// table-of-contents) and guidFileFormat at offset 48 (version) — never the file
/// content or any binary structures. On-disk GUIDs use the standard .NET mixed-endian
/// layout, so <c>new Guid(span)</c> decodes them directly.
/// </summary>
public static class OneFileFormatSniffer
{
    /// <summary>Bytes required to hold both header GUIDs (offset 48 + 16).</summary>
    public const int HeaderLength = 64;

    /// <summary>guidFileType marking a section (.one).</summary>
    public static readonly Guid FileTypeSection = new("7B5C52E4-D88C-4DA7-AEB1-5378D02996D3");

    /// <summary>guidFileType marking a table-of-contents (.onetoc2).</summary>
    public static readonly Guid FileTypeTableOfContents = new("43FF2FA1-EFD9-4C76-9EE2-10EA5722765F");

    /// <summary>guidFileFormat of the documented 2010+ revision-store format.</summary>
    public static readonly Guid FileFormatRevisionStore = new("109ADD3F-911B-49F5-A5D0-1791EDC8AED8");

    /// <summary>
    /// Classifies a OneNote header from its raw bytes. Reads guidFileType at offset 0
    /// to distinguish section/toc, then guidFileFormat at offset 48 to distinguish the
    /// current revision-store format from anything older (treated as legacy/2007).
    /// </summary>
    public static OneFileFormatResult Sniff(ReadOnlySpan<byte> header)
    {
        if (header.Length < HeaderLength)
            return new OneFileFormatResult(
                OneNoteFileFormat.NotOneNoteFile,
                "unknown",
                $"Only {header.Length} bytes available; a OneNote header needs {HeaderLength} bytes " +
                "to hold guidFileType (offset 0) and guidFileFormat (offset 48).");

        var fileType = new Guid(header.Slice(0, 16));
        var fileTypeLabel = LabelFor(fileType);
        if (fileTypeLabel == "unknown")
            return new OneFileFormatResult(
                OneNoteFileFormat.NotOneNoteFile,
                "unknown",
                $"guidFileType {fileType:B} does not match a OneNote section or table-of-contents file.");

        var fileFormat = new Guid(header.Slice(48, 16));
        if (fileFormat == FileFormatRevisionStore)
            return new OneFileFormatResult(
                OneNoteFileFormat.Current2010Plus,
                fileTypeLabel,
                $"guidFileFormat matches the documented 2010+ revision-store format for a {fileTypeLabel} file.");

        return new OneFileFormatResult(
            OneNoteFileFormat.Legacy2007,
            fileTypeLabel,
            $"Valid OneNote {fileTypeLabel} header but guidFileFormat {fileFormat:B} is not the documented " +
            "2010+ revision-store format — treated as legacy/2007; OneNote 2016 may still open & convert it.");
    }

    /// <summary>
    /// Sniffs a file by reading only its leading header bytes. Opens read-only with
    /// FileShare.ReadWrite so it never locks or modifies the file (safe against files
    /// OneNote may hold open). Lets a missing-file error propagate to the caller.
    /// </summary>
    public static OneFileFormatResult SniffFile(string path)
    {
        var buffer = new byte[HeaderLength];
        int bytesRead;
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            bytesRead = ReadUpTo(stream, buffer);
        }

        return Sniff(buffer.AsSpan(0, bytesRead));
    }

    /// <summary>Reads until the buffer is full or the stream ends; returns bytes read.</summary>
    private static int ReadUpTo(Stream stream, byte[] buffer)
    {
        var total = 0;
        while (total < buffer.Length)
        {
            var read = stream.Read(buffer, total, buffer.Length - total);
            if (read == 0)
                break;
            total += read;
        }

        return total;
    }

    /// <summary>Maps a guidFileType to its report label, or "unknown".</summary>
    private static string LabelFor(Guid fileType)
    {
        if (fileType == FileTypeSection)
            return "section";
        if (fileType == FileTypeTableOfContents)
            return "tableOfContents";
        return "unknown";
    }
}
