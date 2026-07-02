using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using OneNoteMcp.Interop;

namespace OneNoteMcp.Tests.Fixtures;

/// <summary>
/// Disposable, clean-room OneNote fixture. Builds a throw-away notebook under
/// %TEMP%\onenote-mcp-tests\{guid} with a known text page and a known image
/// page, so COM integration tests have deterministic content to assert against.
///
/// On machines without OneNote (or where COM fails), <see cref="Available"/>
/// stays false and tests skip gracefully instead of failing the whole run.
/// Disposal closes the notebook and deletes the scratch directory; it never
/// touches any real notebook.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class FixtureNotebook : IDisposable
{
    private const string OneNs = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    // Clean-room 1x1 transparent PNG (generated, not copied from any source notebook).
    private const string OnePixelPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

    /// <summary>True only when the full COM build succeeded and content is ready.</summary>
    public bool Available { get; }

    /// <summary>Scratch directory holding the notebook files.</summary>
    public string Directory { get; } = string.Empty;

    /// <summary>COM object ID of the created notebook.</summary>
    public string NotebookId { get; } = string.Empty;

    /// <summary>COM object ID of the section inside the notebook.</summary>
    public string SectionId { get; } = string.Empty;

    /// <summary>COM object ID of the known text page.</summary>
    public string TextPageId { get; } = string.Empty;

    /// <summary>COM object ID of the known image page.</summary>
    public string ImagePageId { get; } = string.Empty;

    /// <summary>Title placed on the text page.</summary>
    public string KnownTitle => "Fixture Text Page";

    /// <summary>Text run placed on the text page.</summary>
    public string KnownText => "known fixture text run";

    /// <summary>
    /// When the environment variable <c>ONENOTE_COM_REQUIRED=1</c> is set, a
    /// missing or failing COM build throws instead of silently skipping. This is
    /// the guard against false-green runs: on a machine that is SUPPOSED to drive
    /// OneNote live, a dead COM path turns the suite RED rather than passing in
    /// milliseconds without ever exercising COM.
    /// </summary>
    private static bool ComRequired =>
        Environment.GetEnvironmentVariable("ONENOTE_COM_REQUIRED") == "1";

    /// <summary>Builds the fixture. Downgrades to a skip on missing/failing COM.</summary>
    public FixtureNotebook()
    {
        if (!OneNoteSession.IsComAvailable)
        {
            if (ComRequired)
            {
                throw new InvalidOperationException(
                    "ONENOTE_COM_REQUIRED=1 but the OneNote.Application COM ProgID is not " +
                    "registered on this machine. Refusing to skip COM integration tests.");
            }

            // No OneNote on this machine — leave Available = false so tests skip.
            return;
        }

        try
        {
            var session = OneNoteSession.Instance;

            Directory = Path.Combine(
                Path.GetTempPath(), "onenote-mcp-tests", Guid.NewGuid().ToString("N"));
            System.IO.Directory.CreateDirectory(Directory);

            NotebookId = session.OpenHierarchy(Directory, "", OneNoteCreateFileType.CftNotebook);
            SectionId = session.OpenHierarchy("Fixture.one", NotebookId, OneNoteCreateFileType.CftSection);

            TextPageId = session.CreateNewPage(SectionId);
            session.UpdatePageContent(BuildTextPageXml(TextPageId), OneNoteXmlSchema.Xs2013);

            ImagePageId = session.CreateNewPage(SectionId);
            session.UpdatePageContent(BuildImagePageXml(ImagePageId), OneNoteXmlSchema.Xs2013);

            Available = true;
        }
        catch (COMException ex)
        {
            Available = false;
            TryCleanup();

            if (ComRequired)
            {
                // COM is present and required, but the live build failed — fail loud
                // so the false-green can never recur.
                throw new InvalidOperationException(
                    "ONENOTE_COM_REQUIRED=1 but building the fixture notebook over COM " +
                    $"failed (HRESULT 0x{ex.HResult:X8}). See inner exception.", ex);
            }

            // COM present but unusable in this environment — treat as a skip.
        }
    }

    /// <summary>Text page XML: a title plus a single known text run.</summary>
    private string BuildTextPageXml(string pageId) =>
        $"<one:Page xmlns:one=\"{OneNs}\" ID=\"{pageId}\">" +
        $"<one:Title><one:OE><one:T><![CDATA[{KnownTitle}]]></one:T></one:OE></one:Title>" +
        "<one:Outline><one:OEChildren><one:OE>" +
        $"<one:T><![CDATA[{KnownText}]]></one:T>" +
        "</one:OE></one:OEChildren></one:Outline></one:Page>";

    /// <summary>Image page XML: a title plus a single embedded 1x1 PNG.</summary>
    private string BuildImagePageXml(string pageId) =>
        $"<one:Page xmlns:one=\"{OneNs}\" ID=\"{pageId}\">" +
        $"<one:Title><one:OE><one:T><![CDATA[{KnownTitle}]]></one:T></one:OE></one:Title>" +
        "<one:Outline><one:OEChildren><one:OE>" +
        $"<one:Image format=\"png\"><one:Data>{OnePixelPngBase64}</one:Data></one:Image>" +
        "</one:OE></one:OEChildren></one:Outline></one:Page>";

    /// <summary>Closes the notebook and deletes scratch files. Idempotent, never throws.</summary>
    public void Dispose()
    {
        if (Available || !string.IsNullOrEmpty(NotebookId))
        {
            try
            {
                OneNoteSession.Instance.CloseNotebook(NotebookId);
            }
            catch (COMException)
            {
                // Best-effort close — swallow so disposal is never fatal.
            }
        }

        TryCleanup();
    }

    /// <summary>Deletes the scratch directory if present. Swallows all errors.</summary>
    private void TryCleanup()
    {
        try
        {
            if (!string.IsNullOrEmpty(Directory) && System.IO.Directory.Exists(Directory))
            {
                System.IO.Directory.Delete(Directory, recursive: true);
            }
        }
        catch
        {
            // Cleanup is best-effort; leftover temp files are harmless.
        }
    }
}
