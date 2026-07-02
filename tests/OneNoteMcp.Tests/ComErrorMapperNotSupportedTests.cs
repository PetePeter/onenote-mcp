using OneNoteMcp.Interop;
using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Tests that ComErrorMapper.Describe maps NotSupportedInVersionException to a
/// structured, human-readable message mentioning the version display name and method.
/// </summary>
public sealed class ComErrorMapperNotSupportedTests
{
    [Fact]
    public void Describe_NotSupportedInVersion_MentionsVersionAndMethod()
    {
        var ex = new NotSupportedInVersionException(12, "MergeFiles");

        var result = ComErrorMapper.Describe(ex);

        Assert.Contains("OneNote 2007", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("MergeFiles", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Describe_NotSupportedInVersion_ContainsUnsupportedWording()
    {
        var ex = new NotSupportedInVersionException(12, "SyncHierarchy");

        var result = ComErrorMapper.Describe(ex);

        Assert.Contains("Unsupported", result, StringComparison.OrdinalIgnoreCase);
    }
}
