// <copyright file="ResourceController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Configuration
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
                    radiobutton = new
                    {
                        maxRadioChoices = this.localizer.GetString("maxRadioChoices").Value,
                        radioOptions = this.localizer.GetString("radioOptions").Value,
                    },
                    checkbox = new
                    {
                        maxCheckBoxChoices = this.localizer.GetString("maxCheckBoxChoices").Value,
                        checkBoxOptions = this.localizer.GetString("checkBoxOptions").Value,
                    },
                    dropdown = new
                    {
                        dropdownOptions = this.localizer.GetString("dropdownOptions").Value,
                        maxDropdownChoices = this.localizer.GetString("maxDropdownChoices").Value,
                    },
                    common = new
                    {
                        displayName = this.localizer.GetString("displayName").Value,
                        displayNamePlaceholder = this.localizer.GetString("displayNamePlaceholder").Value,
                        placeholder = this.localizer.GetString("placeholder").Value,
                        placeholderText = this.localizer.GetString("placeholderText").Value,
                        btnAddComponent = this.localizer.GetString("btnAddComponent").Value,
                        nonEmptyItem = this.localizer.GetString("nonEmptyItem").Value,
                        duplicateItem = this.localizer.GetString("duplicateItem").Value,
                        minimumItems = this.localizer.GetString("minimumItems").Value,
                        invalidTeamLink = this.localizer.GetString("invalidTeamLink").Value,
                        emptyTeamLink = this.localizer.GetString("emptyTeamLink").Value,
                        successPublish = this.localizer.GetString("successPublish").Value,
                        genericError = this.localizer.GetString("genericError").Value,
                        teamLink = this.localizer.GetString("teamLink").Value,
                        loading = this.localizer.GetString("loading").Value,
                        btnLogout = this.localizer.GetString("btnLogout").Value,
                        mainHeader = this.localizer.GetString("mainHeader").Value,
                        notAuthorized = this.localizer.GetString("notAuthorized").Value,
                        urgentSeverity = this.localizer.GetString("urgentSeverity").Value,
                        normalSeverity = this.localizer.GetString("normalSeverity").Value,
                    },
                    buildForm = new
                    {
                        btnBuildForm = this.localizer.GetString("btnBuildForm").Value,
                        componentDropdown = this.localizer.GetString("componentDropdown").Value,
                        headerTitle = this.localizer.GetString("headerTitle").Value,
                        maxComponents = this.localizer.GetString("maxComponents").Value,
                        maxLengthDisplayName = this.localizer.GetString("maxLengthDisplayName").Value,
                        notEmptyDisplayName = this.localizer.GetString("notEmptyDisplayName").Value,
                        previewTitle = this.localizer.GetString("previewTitle").Value,
                        staticDropdown = this.localizer.GetString("staticDropdown").Value,
                        staticDropdownPlaceholder = this.localizer.GetString("staticDropdownPlaceholder").Value,
                        titleText = this.localizer.GetString("titleText").Value,
                        descriptionPlaceholderText = this.localizer.GetString("descriptionPlaceholderText").Value,
                        descriptionText = this.localizer.GetString("descriptionText").Value,
                        uniqueDisplayName = this.localizer.GetString("uniqueDisplayName").Value,
                        checkBox = this.localizer.GetString("checkBox").Value,
                        dropDown = this.localizer.GetString("dropDown").Value,
                        inputDate = this.localizer.GetString("inputDate").Value,
                        inputText = this.localizer.GetString("inputText").Value,
                        radioButton = this.localizer.GetString("radioButton").Value,
                    },
                };
                return this.Ok(strings);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching resource strings.");
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }
    }
}