// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Configuration
{
    using Microsoft.AspNetCore;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Program main class.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Main method.
        /// </summary>
        /// <param name="args">string of arguments.</param>
        public static void Main(string[] args)
        {
            CreateWebHostBuilder(args).Build().Run();
        }

        /// <summary>
        /// This method sets up configurations of the application.
        /// </summary>
        /// <param name="args">string of arguments.</param>
        /// <returns>unit of execution.</returns>
        public static IWebHostBuilder CreateWebHostBuilder(string[] args) =>
            WebHost.CreateDefaultBuilder(args)
                .ConfigureLogging((hostContext, logging) =>
                {
                    var appInsightKey = hostContext.Configuration.GetSection("ApplicationInsights")["InstrumentationKey"];
                    logging.AddApplicationInsights(appInsightKey);
                })
                .UseStartup<Startup>();
    }
}
