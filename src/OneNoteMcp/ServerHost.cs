using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace OneNoteMcp;

/// <summary>
/// Builds the generic host that runs the OneNote MCP server over stdio.
/// Extracted from <c>Program</c> so tests can build the host without starting
/// the stdio transport.
/// </summary>
public static class ServerHost
{
    /// <summary>
    /// Creates a configured host builder: MCP server, stdio transport, and all
    /// tools discovered by attribute in this assembly.
    /// </summary>
    public static IHostBuilder CreateBuilder(string[] args)
    {
        return Host.CreateDefaultBuilder(args)
            .ConfigureLogging(logging =>
            {
                // stdout is reserved for the MCP protocol; route logs to stderr only.
                logging.ClearProviders();
                logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);
            })
            .ConfigureServices(services =>
            {
                services
                    .AddMcpServer()
                    .WithStdioServerTransport()
                    .WithToolsFromAssembly(typeof(ServerHost).Assembly);
            });
    }
}
