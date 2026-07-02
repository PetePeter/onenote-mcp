using OneNoteMcp.Interop;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>Tests for the typed "method missing on this version" exception.</summary>
public sealed class NotSupportedInVersionExceptionTests
{
    [Fact]
    public void CarriesVersionAndMethod()
    {
        var ex = new NotSupportedInVersionException(12, "MergeSections");
        Assert.Equal(12, ex.VersionMajor);
        Assert.Equal("MergeSections", ex.MethodName);
    }

    [Fact]
    public void Message_MentionsVersionAndMethod()
    {
        var ex = new NotSupportedInVersionException(12, "MergeSections");
        Assert.Contains("12", ex.Message);
        Assert.Contains("MergeSections", ex.Message);
    }

    [Fact]
    public void IsNotSupportedException()
    {
        var ex = new NotSupportedInVersionException(12, "SyncHierarchy");
        Assert.IsAssignableFrom<System.NotSupportedException>(ex);
    }
}
