using Xunit;

namespace OneNoteMcp.Tests;

/// <summary>
/// Ensures all tests in the "OneNote COM" collection run sequentially,
/// preventing concurrent access to the single-threaded COM object.
/// </summary>
[CollectionDefinition("OneNote COM", DisableParallelization = true)]
public sealed class OneNoteComCollection { }
