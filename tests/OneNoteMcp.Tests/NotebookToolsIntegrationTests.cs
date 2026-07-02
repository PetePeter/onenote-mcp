using System;
using System.IO;
using System.Linq;
using OneNoteMcp.Interop;
using OneNoteMcp.Model;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// End-to-end COM integration test for the notebook write tools. Uses its OWN
/// scratch notebook folder under %TEMP% and closes + deletes it on the way out,
/// so it never touches the shared fixture or any real notebook. Skips when the
/// fixture (and thus live COM) is unavailable.
/// </summary>
[Collection("OneNote COM")]
public sealed class NotebookToolsIntegrationTests
{
    private readonly FixtureNotebook _fx;

    public NotebookToolsIntegrationTests(FixtureNotebook fx) => _fx = fx;

    [Fact]
    public void CreateCloseOpen_ScratchNotebook_TogglesInListNotebooks()
    {
        if (!_fx.Available) return;

        var scratchDir = Path.Combine(
            Path.GetTempPath(), "onenote-mcp-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(scratchDir);

        string? id = null;
        try
        {
            id = NotebookTools.CreateNotebook(scratchDir);

            Assert.Contains(OpenNotebooks(), n => n.Id == id); // present after create

            NotebookTools.CloseNotebook(id);

            Assert.DoesNotContain(OpenNotebooks(), n => n.Id == id); // gone after close
        }
        finally
        {
            if (id is not null)
            {
                try { NotebookTools.CloseNotebook(id); } catch { /* already closed */ }
            }
            try { Directory.Delete(scratchDir, recursive: true); } catch { /* best-effort */ }
        }
    }

    /// <summary>Reads the currently open notebooks from live hierarchy XML.</summary>
    private static System.Collections.Generic.IReadOnlyList<NotebookInfo> OpenNotebooks()
    {
        var xml = OneNoteSession.Instance.GetHierarchy(
            "", OneNoteScope.HsNotebooks, OneNoteXmlSchema.Xs2013);
        return HierarchyParser.ParseNotebooks(xml);
    }
}
