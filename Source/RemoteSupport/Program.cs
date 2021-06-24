// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport
{
    using Microsoft.AspNetCore;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// The Program class is responsible for holding the entry-point of the program.
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// The entry-point for the program.
        /// </summary>
        /// <param name="args">The command line arguments.</param>
        public static void Main(string[] args)
        {
            CreateWebHostBuilder(args).Build().Run();
        }

        /// <summary>
        /// Build the web-host for servicing HTTP requests.
        /// </summary>
        /// <param name="args">The command line arguments.</param>
        /// <returns> The WebHostBuilder configured from the arguments with the composition root defined in <see cref="Startup" />.</returns>
        public static IWebHostBuilder CreateWebHostBuilder(string[] args) =>
            WebHost
                .CreateDefaultBuilder(args)
                .ConfigureLogging((hostingContext, logging) =>
                {
                    // hostingContext.HostingEnvironment can be used to determine environments as well.
                    var appInsightKey = hostingContext.Configuration.GetSection("ApplicationInsights")["InstrumentationKey"];
                    logging.AddApplicationInsights(appInsightKey);
                })
                .ConfigureAppConfiguration((hostingContext, config) =>
                {
                    config
                        .AddEnvironmentVariables();

                    if (hostingContext.HostingEnvironment.IsDevelopment())
                    {
                        // Using dot net secrets to store the settings during development
                        // https://docs.microsoft.com/en-us/aspnet/core/security/app-secrets?view=aspnetcore-3.0&tabs=windows
                        config.AddUserSecrets<Startup>();
                    }
                })
                .UseStartup<Startup>();
    }
}
