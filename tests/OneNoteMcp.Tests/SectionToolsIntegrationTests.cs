using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OneNoteMcp.Interop;
using OneNoteMcp.Tests.Fixtures;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// End-to-end COM integration tests for the section write tools (create / rename
/// / delete). Each test creates its OWN scratch section under the shared fixture
/// notebook and deletes it on the way out, so it never mutates content other
/// tests assert against. They early-return (skip) when the fixture is unavailable.
/// </summary>
[Collection("OneNote COM")]
public sealed class SectionToolsIntegrationTests
{
    private readonly FixtureNotebook _fx;

    public SectionToolsIntegrationTests(FixtureNotebook fx) => _fx = fx;

    [Fact]
    public void CreateSection_UnderFixtureNotebook_AppearsInHierarchy()
    {
        if (!_fx.Available) return;

        var name = "Scratch Section " + Guid.NewGuid().ToString("N");
        var id = SectionTools.CreateSection(_fx.NotebookId, name);
        try
        {
            Assert.False(string.IsNullOrEmpty(id));
            Assert.Contains(NotebookSections(), s => s.Id == id && s.Name == name);
        }
        finally
        {
            SectionTools.DeleteNode(id);
        }
    }

    [Fact]
    public void RenameNode_ChangesSectionNameInHierarchy()
    {
        if (!_fx.Available) return;

        var id = SectionTools.CreateSection(_fx.NotebookId, "Scratch Section " + Guid.NewGuid().ToString("N"));
        try
        {
            var newName = "Renamed " + Guid.NewGuid().ToString("N");
            SectionTools.RenameNode(id, newName);

            Assert.Contains(NotebookSections(), s => s.Id == id && s.Name == newName);
        }
        finally
        {
            SectionTools.DeleteNode(id);
        }
    }

    [Fact]
    public void DeleteNode_RemovesSectionFromHierarchy()
    {
        if (!_fx.Available) return;

        var id = SectionTools.CreateSection(_fx.NotebookId, "Scratch Section " + Guid.NewGuid().ToString("N"));
        try
        {
            Assert.Contains(NotebookSections(), s => s.Id == id); // present after create

            SectionTools.DeleteNode(id);

            Assert.DoesNotContain(NotebookSections(), s => s.Id == id); // gone after delete
        }
        finally
        {
            // Best-effort cleanup if an assertion above aborted before the delete.
            try { SectionTools.DeleteNode(id); } catch { /* already deleted */ }
        }
    }

    [Fact]
    public void DeleteNode_NonExistentId_ReturnsMappedErrorWithoutThrowing()
    {
        if (!_fx.Available) return;

        // A well-formed but non-existent object ID must NOT throw: the tool routes
        // the COM failure through the central mapper and returns a human-readable
        // error string as its result (errors-in-tool-results design).
        var result = SectionTools.DeleteNode("{00000000-0000-0000-0000-000000000000}{1}{B0}");

        Assert.Contains("OneNote error", result, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("\"deleted\":true", result); // not a success payload
    }

    /// <summary>Reads the current sections under the fixture notebook from live hierarchy XML.</summary>
    private IReadOnlyList<(string Id, string Name)> NotebookSections()
    {
        var xml = OneNoteSession.Instance.GetHierarchy(
            _fx.NotebookId, OneNoteScope.HsSections, OneNoteXmlSchema.Xs2013);

        var root = XDocument.Parse(xml).Root;
        if (root is null) return Array.Empty<(string, string)>();

        return root.DescendantsAndSelf()
            .Where(e => e.Name.LocalName == "Section")
            .Select(e => (
                Id: e.Attribute("ID")?.Value ?? "",
                Name: e.Attribute("name")?.Value ?? ""))
            .ToList();
    }
}
