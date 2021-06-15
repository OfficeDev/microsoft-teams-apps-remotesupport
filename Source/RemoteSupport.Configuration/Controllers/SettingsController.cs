// <copyright file="SettingsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Configuration
{
    using System;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Configuration.Models;

    /// <summary>
    /// This endpoint is used to provide app settings to client app.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class SettingsController : ControllerBase
    {
        private readonly AzureAdSettings azureAdSettings;
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="SettingsController"/> class.
        /// </summary>
        /// <param name="options">Azure ad configurations.</param>
        /// <param name="logger">Logger instance.</param>
        public SettingsController(IOptionsMonitor<AzureAdSettings> options, ILogger<SettingsController> logger)
        {
            this.azureAdSettings = options?.CurrentValue;
            this.logger = logger;
        }

        /// <summary>
        /// This endpoint gives Azure AD configurations.
        /// </summary>
        /// <returns>object.</returns>
        [HttpGet]
        public ActionResult GetSettings()
        {
            try
            {
                return this.Ok(new
                {
                    ClientId = this.azureAdSettings.ClientId,
                    TenantId = this.azureAdSettings.Tenant,
                    TokenEndpoint = this.azureAdSettings.Instance + this.azureAdSettings.Tenant,
                });
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in fetching app settings");
                throw;
            }
        }
    }
}