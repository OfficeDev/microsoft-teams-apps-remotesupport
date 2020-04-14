// <copyright file="ResourceController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Controllers
{
    using System;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Controller to handle resource strings related request.
    /// </summary>
    [Route("api/resource")]
    [Authorize]
    [ApiController]
    public class ResourceController : ControllerBase
    {
        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResourceController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        public ResourceController(ILogger<ResourceController> logger, IStringLocalizer<Strings> localizer)
        {
            this.logger = logger;
            this.localizer = localizer;
        }

        /// <summary>
        /// Get resource strings for displaying in client application.
        /// </summary>
        /// <returns>Resource strings according to user locale.</returns>
        [Route("resourcestrings")]
        public IActionResult GetResourceStrings()
        {
            try
            {
                var strings = new
                {
                    AddButtonText = this.localizer.GetString("AddButtonText").Value,
                    SaveButtonText = this.localizer.GetString("SaveButtonText").Value,
                    ExpertListTitle = this.localizer.GetString("ExpertListTitle").Value,
                    ExpertNameTitle = this.localizer.GetString("ExpertNameTitle").Value,
                    NoMatchesFoundText = this.localizer.GetString("NoMatchesFoundText").Value,
                    ExpertListPlaceHolderText = this.localizer.GetString("ExpertListPlaceHolderText").Value,
                    ErrorMessage = this.localizer.GetString("ErrorMessage").Value,
                    MaxOnCallExpertsAllowedText = this.localizer.GetString("MaxOnCallExpertsAllowedText").Value,
                    UnauthorizedAccess = this.localizer.GetString("UnauthorizedAccess").Value,
                    SessionExpired = this.localizer.GetString("SessionExpired").Value,
                };
                return this.Ok(strings);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while fetching resource strings.");
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }
    }
}