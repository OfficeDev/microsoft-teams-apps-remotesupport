// <copyright file="StorageController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Configuration
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Common;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Providers;
    using Microsoft.Teams.Apps.RemoteSupport.Configuration.Models;

    /// <summary>
    /// Storage API provider class.
    /// </summary>
    [Route("api/[controller]")]
    [Authorize]
    [ApiController]
    public class StorageController : ControllerBase
    {
        private readonly ICardConfigurationStorageProvider configurationStorageProvider;
        private readonly IEnumerable<string> validUsers;
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="StorageController"/> class.
        /// </summary>
        /// <param name="configurationStorageProvider">Configuration storage provider.</param>
        /// <param name="options">Azure Active Directory Configurations.</param>
        /// <param name="logger">Logger instance.</param>
        public StorageController(ICardConfigurationStorageProvider configurationStorageProvider, IOptionsMonitor<AzureAdSettings> options, ILogger<StorageController> logger)
        {
            this.configurationStorageProvider = configurationStorageProvider;
            var azureAdSettings = options?.CurrentValue;
            this.validUsers = azureAdSettings.Upn.Split(';');
            this.logger = logger;
        }

        /// <summary>
        /// This endpoint is used to store adaptive card json configuration details in storage.
        /// </summary>
        /// <param name="configurationEntity">Configuration details.</param>
        /// <returns>Task.</returns>
        [HttpPost]
        public async Task<IActionResult> SaveConfigurationsAsync([FromBody]CardConfigurationEntity configurationEntity)
        {
            try
            {
                string user = this.HttpContext.User.Identity.Name;
                if (!this.validUsers.Contains(user))
                {
                    return this.Unauthorized();
                }

                configurationEntity.CardId = Guid.NewGuid().ToString();
                configurationEntity.CreatedOn = DateTime.UtcNow;
                configurationEntity.CreatedByUserPrincipalName = user;
                configurationEntity.CreatedByObjectId = this.GetId();
                configurationEntity.TeamId = Utility.ParseTeamIdFromDeepLink(configurationEntity.TeamLink);

                var result = await this.configurationStorageProvider.StoreOrUpdateEntityAsync(configurationEntity);
                if (result == null)
                {
                    this.logger.LogInformation("Error in saving configurations " + configurationEntity.CreatedByObjectId);
                    return this.StatusCode(StatusCodes.Status500InternalServerError);
                }

                this.logger.LogInformation("Configurations saved successfully " + configurationEntity.CreatedByObjectId);
                return this.Ok();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in saving configurations");
                throw;
            }
        }

        /// <summary>
        /// This method is used to get the latest configuration details.
        /// </summary>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        [HttpGet]
        public async Task<IActionResult> GetConfigurationsAsync()
        {
            try
            {
                if (!this.validUsers.Contains(this.HttpContext.User.Identity.Name))
                {
                    this.logger.LogInformation("Unauthorized " + this.GetId());
                    return this.Unauthorized();
                }

                this.logger.LogInformation("Get configurations " + this.GetId());
                var result = await this.configurationStorageProvider.GetConfigurationAsync();
                return this.Ok(result);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in getting configurations");
                throw;
            }
        }

        /// <summary>
        /// Get the id of the user in Azure AD (GUID format).
        /// </summary>
        /// <returns> Returns the id of the user in Azure AD (GUID format). </returns>
        private string GetId()
        {
            var idClaims = this.HttpContext.User.Claims
                .FirstOrDefault(claim => claim.Type == AzureAdClaimTypes.ObjectId);

            return idClaims?.Value;
        }
    }
}
