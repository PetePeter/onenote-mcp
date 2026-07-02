using System;
using System.Runtime.InteropServices;
using OneNoteMcp.Interop;
using OneNoteMcp.Tools;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Pure unit tests for <see cref="ToolError"/> — the guard that every COM-touching
/// tool runs its body through so a failure becomes a human-readable tool result
/// instead of an unhandled exception carrying a raw HRESULT.
/// </summary>
public sealed class ToolErrorTests
{
    private const int HrObjectDoesNotExist = unchecked((int)0x80042014);

    [Fact]
    public void Guard_Success_ReturnsBodyResult()
    {
        Assert.Equal("ok", ToolError.Guard(() => "ok"));
    }

    [Fact]
    public void Guard_ComException_ReturnsMappedString()
    {
        var ex = new COMException("raw", HrObjectDoesNotExist);
        var result = ToolError.Guard(() => throw ex);

        Assert.Equal(ComErrorMapper.Describe(ex), result);
    }

    [Fact]
    public void Guard_NonComException_ReturnsReadableMessage()
    {
        var result = ToolError.Guard(() => throw new InvalidOperationException("bad-arg-detail"));
        Assert.Contains("bad-arg-detail", result);
    }
}
