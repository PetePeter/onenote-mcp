using Microsoft.Extensions.Hosting;
using OneNoteMcp;

await ServerHost.CreateBuilder(args).Build().RunAsync();
