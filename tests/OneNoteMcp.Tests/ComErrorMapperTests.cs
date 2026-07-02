using System;
using System.Runtime.InteropServices;
using OneNoteMcp.Interop;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Pure, COM-free unit tests for <see cref="ComErrorMapper"/> — the central
/// translator from raw OneNote/COM HRESULTs to human-readable messages plus a
/// suggested next action. No OneNote runtime is required.
/// </summary>
public sealed class ComErrorMapperTests
{
    private const int HrObjectDoesNotExist = unchecked((int)0x80042014);
    private const int HrLegacySection      = unchecked((int)0x8004201E);
    private const int HrAppInModalUI       = unchecked((int)0x80042030);
    private const int HrServerUnavailable  = unchecked((int)0x800706BA);
    private const int HrUnknown            = unchecked((int)0x80041234);

    [Fact]
    public void Lookup_KnownCode_ReturnsInfoWithSymbolMessageAndAction()
    {
        var info = ComErrorMapper.Lookup(HrObjectDoesNotExist);

        Assert.NotNull(info);
        Assert.Equal(HrObjectDoesNotExist, info!.HResult);
        Assert.False(string.IsNullOrWhiteSpace(info.Symbol));
        Assert.False(string.IsNullOrWhiteSpace(info.Message));
        Assert.False(string.IsNullOrWhiteSpace(info.SuggestedAction));
    }

    [Fact]
    public void Lookup_UnknownCode_ReturnsNullAndIsKnownFalse()
    {
        Assert.Null(ComErrorMapper.Lookup(HrUnknown));
        Assert.False(ComErrorMapper.IsKnown(HrUnknown));
        Assert.True(ComErrorMapper.IsKnown(HrObjectDoesNotExist));
    }

    [Fact]
    public void Describe_Int_KnownCode_IncludesHexMessageAndAction()
    {
        var info = ComErrorMapper.Lookup(HrObjectDoesNotExist)!;
        var text = ComErrorMapper.Describe(HrObjectDoesNotExist);

        Assert.Contains("0x80042014", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(info.Message, text);
        Assert.Contains(info.SuggestedAction, text);
    }

    [Fact]
    public void Describe_Int_UnknownCode_StillIncludesHexCode()
    {
        var text = ComErrorMapper.Describe(HrUnknown);
        Assert.Contains("0x80041234", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Describe_Exception_ComException_MapsByHResult()
    {
        var ex = new COMException("raw", HrObjectDoesNotExist);
        Assert.Equal(ComErrorMapper.Describe(HrObjectDoesNotExist), ComErrorMapper.Describe(ex));
    }

    [Fact]
    public void Describe_Exception_UnknownComHResult_IncludesHexCode()
    {
        var ex = new COMException("raw", HrUnknown);
        Assert.Contains("0x80041234", ComErrorMapper.Describe(ex), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Describe_Exception_NonComException_ReturnsMessage()
    {
        var text = ComErrorMapper.Describe(new InvalidOperationException("boom-detail"));
        Assert.Contains("boom-detail", text);
    }

    [Fact]
    public void Describe_KnownCodes_HaveCodeSpecificDistinctActions()
    {
        // Guards against a mapper that returns one generic action for every code.
        var objAction    = ComErrorMapper.Lookup(HrObjectDoesNotExist)!.SuggestedAction;
        var legacyAction = ComErrorMapper.Lookup(HrLegacySection)!.SuggestedAction;
        var modalAction  = ComErrorMapper.Lookup(HrAppInModalUI)!.SuggestedAction;

        Assert.NotEqual(objAction, legacyAction);
        Assert.NotEqual(objAction, modalAction);
        Assert.NotEqual(legacyAction, modalAction);

        // And each carries its distinctive, actionable hint.
        Assert.Contains("onenote_list_notebooks", objAction, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("onenote_convert_section", legacyAction, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("dialog", modalAction, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Describe_Int_ServerUnavailable_IsMappedNotRaw()
    {
        // Infra RPC code must also be catalogued (not fall through to unknown).
        Assert.True(ComErrorMapper.IsKnown(HrServerUnavailable));
        Assert.Contains("0x800706BA", ComErrorMapper.Describe(HrServerUnavailable), StringComparison.OrdinalIgnoreCase);
    }
}
