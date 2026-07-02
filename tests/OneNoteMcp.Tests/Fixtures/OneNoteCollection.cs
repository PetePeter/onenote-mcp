using Xunit;

namespace OneNoteMcp.Tests.Fixtures;

/// <summary>
/// The single serialized COM collection. DisableParallelization keeps all
/// OneNote COM tests on one thread, and ICollectionFixture shares one built
/// <see cref="FixtureNotebook"/> across every test in the collection.
/// The name "OneNote COM" is load-bearing: existing [Collection("OneNote COM")]
/// attributes bind to it.
/// </summary>
[CollectionDefinition("OneNote COM", DisableParallelization = true)]
public sealed class OneNoteCollection : ICollectionFixture<FixtureNotebook> { }
