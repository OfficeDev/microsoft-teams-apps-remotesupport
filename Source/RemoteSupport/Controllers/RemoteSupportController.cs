// <copyright file="RemoteSupportController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Providers;

    /// <summary>
    /// Controller to handle Remote Support API operations.
    /// </summary>
    [Route("api/remotesupport")]
    [ApiController]
    [Authorize]
    public class RemoteSupportController : BaseRemoteSupportController
    {
        /// <summary>
        /// Microsoft Application ID.
        /// </summary>
        private readonly string appId;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Remote Support Bot adapter to get context.
        /// </summary>
        private readonly BotFrameworkAdapter botAdapter;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Remote Support bot.
        /// </summary>
        private readonly IOptions<RemoteSupportActivityHandlerOptions> options;

        /// <summary>
        /// Provider to search on call support details in Azure Table Storage.
        /// </summary>
        private readonly IOnCallSupportDetailSearchService onCallSupportDetailSearchService;

        /// <summary>
        /// Provider to store on call support details to Azure Table Storage.
        /// </summary>
        private readonly IOnCallSupportDetailStorageProvider onCallSupportDetailStorageProvider;

        /// <summary>
        /// Provider to store card configuration details in Azure Table Storage.
        /// </summary>
        private readonly ICardConfigurationStorageProvider cardConfigurationStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="RemoteSupportController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="botAdapter">Remote support bot adapter.</param>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="onCallSupportDetailSearchService">Provider to search on call support details in Azure Table Storage.</param>
        /// <param name="onCallSupportDetailStorageProvider">Provider to store on call support details in Azure Table Storage.</param>
        /// <param name="cardConfigurationStorageProvider">Provider to store card configuration details in Azure Table Storage.</param>
        public RemoteSupportController(
            ILogger<RemoteSupportController> logger,
            BotFrameworkAdapter botAdapter,
            MicrosoftAppCredentials microsoftAppCredentials,
            IOptions<RemoteSupportActivityHandlerOptions> options,
            IOnCallSupportDetailSearchService onCallSupportDetailSearchService,
            IOnCallSupportDetailStorageProvider onCallSupportDetailStorageProvider,
            ICardConfigurationStorageProvider cardConfigurationStorageProvider)
            : base()
        {
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.logger = logger;
            this.botAdapter = botAdapter;
            this.appId = microsoftAppCredentials != null ? microsoftAppCredentials.MicrosoftAppId : throw new ArgumentNullException(nameof(microsoftAppCredentials));
            this.onCallSupportDetailSearchService = onCallSupportDetailSearchService;
            this.onCallSupportDetailStorageProvider = onCallSupportDetailStorageProvider;
            this.cardConfigurationStorageProvider = cardConfigurationStorageProvider;
        }

        /// <summary>
        /// Get list of members present in a team.
        /// </summary>
        /// <returns>List of members in team.</returns>
        [Route("teammembers")]
        public async Task<IActionResult> GetTeamMembersAsync()
        {
            try
            {
                var cardConfigurationEntity = await this.cardConfigurationStorageProvider?.GetConfigurationAsync();
                string teamId = (cardConfigurationEntity != null) ? cardConfigurationEntity.TeamId : this.options.Value.TeamId;

                var userClaims = this.GetUserClaims();

                var teamsChannelAccounts = new List<TeamsChannelAccount>();
                var conversationReference = new ConversationReference
                {
                    ChannelId = teamId,
                    ServiceUrl = userClaims.ApplicationBasePath,
                };

                await this.botAdapter.ContinueConversationAsync(
                    this.appId,
                    conversationReference,
                    async (context, token) =>
                    {
                        string continuationToken = null;
                        do
                        {
                            var currentPage = await TeamsInfo.GetPagedTeamMembersAsync(context, teamId, continuationToken, pageSize: 500, token);
                            continuationToken = currentPage.ContinuationToken;
                            teamsChannelAccounts.AddRange(currentPage.Members);
                        }
                        while (continuationToken != null);
                    }, default);

                this.logger.LogInformation("GET call for fetching team members from team roster is successful.");
                return this.Ok(teamsChannelAccounts.Select(member => new { content = member.Email, header = member.Name, aadobjectid = member.AadObjectId }));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error occurred while getting team member list.");
                throw;
            }
        }

        /// <summary>
        /// Get list of on call experts configured.
        /// </summary>
        /// <returns>List of on call experts in team.</returns>
        [Route("oncallexperts")]
        public async Task<IActionResult> GetOnCallExpertsAsync()
        {
            try
            {
                if (!this.IsUserAuthenticated())
                {
                    throw new UnauthorizedAccessException("Failed to get fromId from token.");
                }

                this.logger.LogInformation("Initiated call to on call support search service");
                var onCallSupportDetails = await this.onCallSupportDetailSearchService.SearchOnCallSupportTeamAsync(searchQuery: string.Empty, count: 1);
                this.logger.LogInformation("GET call for fetching on call support details is successful.");

                return this.Ok(onCallSupportDetails);
            }
            catch (UnauthorizedAccessException ex)
            {
                this.logger.LogError(ex, "Failed to get user token to make GET call to API.");
                return this.GetErrorResponse(StatusCodes.Status401Unauthorized, ex.Message);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while getting on call experts list.");
                throw;
            }
        }

        /// <summary>
        /// Post call to save on call expert list in Azure Table storage.
        /// </summary>
        /// <param name="onCallSupportDetails">Class contains details of on call support team.</param>
        /// <returns>Returns true for successful operation.</returns>
        [Route("saveoncallsupportdetails")]
        [HttpPost]
        public async Task<IActionResult> SaveOnCallSupportDetailsAsync([FromBody]OnCallSupportDetail onCallSupportDetails)
        {
            try
            {
                if (onCallSupportDetails == null)
                {
                    return this.BadRequest();
                }

                if (!this.IsUserAuthenticated())
                {
                    throw new UnauthorizedAccessException("Failed to get fromId from token.");
                }

                this.logger.LogInformation("Initiated call to on storage provider service.");
                var result = await this.onCallSupportDetailStorageProvider.UpsertOnCallSupportDetailsAsync(onCallSupportDetails);
                this.logger.LogInformation("POST call for saving on call support details in storage is successful.");
                return this.Ok(result);
            }
            catch (UnauthorizedAccessException ex)
            {
                this.logger.LogError(ex, "Failed to get user token to make POST call to API.");
                return this.GetErrorResponse(StatusCodes.Status401Unauthorized, ex.Message);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                this.logger.LogError(ex, "Error while saving on call support details.");
                throw;
            }
        }
    }
}
