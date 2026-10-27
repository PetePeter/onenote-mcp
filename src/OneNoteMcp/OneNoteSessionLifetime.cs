using Microsoft.Extensions.Hosting;
using OneNoteMcp.Interop;

namespace OneNoteMcp;

/// <summary>
/// Disposes every cached <see cref="OneNoteSession"/> when the host stops, so the
/// OneNote COM reference is released deterministically and OneNote can exit.
/// </summary>
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
public sealed class OneNoteSessionLifetime : IHostedService
{
    public Task StartAsync(CancellationToken cancellationToken) => Task.CompletedTask;

    public Task StopAsync(CancellationToken cancellationToken)
    {
        OneNoteSession.DisposeAll();
        return Task.CompletedTask;
    }
}
