using System.Runtime.InteropServices;
using System.Xml.Linq;
using OneNoteMcp.Interop;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Integration tests for OneNoteSession.
/// These tests are READ-ONLY: they only call GetHierarchy and inspect the result.
/// They skip (via early return) when OneNote is not installed or its COM type
/// library is not fully registered on the machine.
/// They will NEVER create, modify, or delete any real notebooks.
/// </summary>
[Collection("OneNote COM")]
public sealed class OneNoteSessionIntegrationTests
{
    [Fact]
    public void GetHierarchy_ReturnsValidNotebooksXml()
    {
        if (!OneNoteSession.IsComAvailable)
        {
            // OneNote ProgID not registered — skip gracefully
            return;
        }

        string xml;
        try
        {
            xml = OneNoteSession.Instance.GetHierarchy(
                startNodeId: "",
                scope: OneNoteScope.HsNotebooks,
                xmlSchema: OneNoteXmlSchema.Xs2013);
        }
        catch (COMException)
        {
            // OneNote ProgID exists but the type library or COM server is not
            // usable in this environment (e.g. TYPE_E_LIBNOTREGISTERED on a
            // machine where Office was installed without full registration).
            // Treat as a skip rather than a failure.
            return;
        }

        Assert.False(string.IsNullOrWhiteSpace(xml), "GetHierarchy returned empty XML");

        var doc = XDocument.Parse(xml); // throws if not well-formed
        Assert.Equal("Notebooks", doc.Root!.Name.LocalName);
    }

    [Fact]
    public void DetectedVersionDisplay_IsNonEmptyWhenComAvailable()
    {
        if (!OneNoteSession.IsComAvailable)
        {
            return;
        }

        var display = OneNoteSession.Instance.DetectedVersionDisplay;
        Assert.False(string.IsNullOrWhiteSpace(display),
            "DetectedVersionDisplay should not be empty when OneNote is installed");
    }
}
