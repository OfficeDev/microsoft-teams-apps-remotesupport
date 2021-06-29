// <copyright file="RemoteSupportActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Cards;
    using Microsoft.Teams.Apps.RemoteSupport.Common;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Providers;
    using Microsoft.Teams.Apps.RemoteSupport.Helpers;
    using Microsoft.Teams.Apps.RemoteSupport.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// The RemoteSupportActivityHandler is responsible for reacting to incoming events from Teams sent from BotFramework.
    /// </summary>
    public sealed class RemoteSupportActivityHandler : TeamsActivityHandler
    {
        /// <summary>
        /// Represents the conversation type as personal.
        /// </summary>
        private const string PersonalConversationType = "personal";

        /// <summary>
        ///  Represents the conversation type as channel.
        /// </summary>
        private const string ChannelConversationType = "channel";

        /// <summary>
        /// Represents the Application base Uri.
        /// </summary>
        private readonly string appBaseUrl;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Remote Support bot.
        /// </summary>
        private readonly IOptions<RemoteSupportActivityHandlerOptions> options;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<RemoteSupportActivityHandler> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// The Application Insights telemetry client.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Microsoft Application credentials for Bot/ME.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Provider to store ticket details to Azure Table Storage.
        /// </summary>
        private readonly ITicketDetailStorageProvider ticketDetailStorageProvider;

        /// <summary>
        /// Provider to search ticket details in Azure Table Storage.
        /// </summary>
        private readonly ITicketSearchService ticketSearchService;

        /// <summary>
        /// Provider to search card configuration details in Azure Table Storage.
        /// </summary>
        private readonly ICardConfigurationStorageProvider cardConfigurationStorageProvider;

        /// <summary>
        /// Provider to search on call support details in Azure Table Storage.
        /// </summary>
        private readonly IOnCallSupportDetailSearchService onCallSupportDetailSearchService;

        /// <summary>
        /// Provider for generating ticket id from Azure Table Storage.
        /// </summary>
        private readonly ITicketIdGeneratorStorageProvider ticketGenerateStorageProvider;

        /// <summary>
        /// Provider for fetching and storing information about on call support in storage table.
        /// </summary>
        private readonly IOnCallSupportDetailStorageProvider onCallSupportDetailStorageProvider;

        /// <summary>
        /// Generating custom JWT token and retrieving access token for user.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Represents unique id of a Team.
        /// </summary>
        private readonly string teamId;

        /// <summary>
        /// Cache for storing objectId's of on call experts.
        /// </summary>
        private readonly IMemoryCache memoryCache;

        /// <summary>
        /// Initializes a new instance of the <see cref="RemoteSupportActivityHandler"/> class.
        /// </summary>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client. </param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="ticketDetailStorageProvider">Provider to store ticket details to Azure Table Storage.</param>
        /// <param name="onCallSupportDetailSearchService">Provider to search on call support details in Azure Table Storage.</param>
        /// <param name="ticketSearchService">Provider to search ticket details in Azure Table Storage.</param>
        /// <param name="tokenHelper">Generating custom JWT token and retrieving access token for user.</param>
        /// <param name="cardConfigurationStorageProvider">Provider to search card configuration details in Azure Table Storage.</param>
        /// <param name="ticketGenerateStorageProvider">Provider to get ticket id to Azure Table Storage.</param>
        /// <param name="onCallSupportDetailStorageProvider"> Provider for fetching and storing information about on call support in storage table.</param>
        /// <param name="memoryCache">MemoryCache instance for caching on call expert objectId's.</param>
        public RemoteSupportActivityHandler(
            MicrosoftAppCredentials microsoftAppCredentials,
            ILogger<RemoteSupportActivityHandler> logger,
            IStringLocalizer<Strings> localizer,
            TelemetryClient telemetryClient,
            IOptions<RemoteSupportActivityHandlerOptions> options,
            ITicketDetailStorageProvider ticketDetailStorageProvider,
            IOnCallSupportDetailSearchService onCallSupportDetailSearchService,
            ITicketSearchService ticketSearchService,
            ICardConfigurationStorageProvider cardConfigurationStorageProvider,
            ITokenHelper tokenHelper,
            ITicketIdGeneratorStorageProvider ticketGenerateStorageProvider,
            IOnCallSupportDetailStorageProvider onCallSupportDetailStorageProvider,
            IMemoryCache memoryCache)
        {
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.logger = logger;
            this.localizer = localizer;
            this.telemetryClient = telemetryClient;
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.teamId = this.options.Value.TeamId;
            this.ticketDetailStorageProvider = ticketDetailStorageProvider;
            this.ticketSearchService = ticketSearchService;
            this.onCallSupportDetailSearchService = onCallSupportDetailSearchService;
            this.appBaseUrl = this.options.Value.AppBaseUri;
            this.tokenHelper = tokenHelper;
            this.cardConfigurationStorageProvider = cardConfigurationStorageProvider;
            this.ticketGenerateStorageProvider = ticketGenerateStorageProvider;
            this.onCallSupportDetailStorageProvider = onCallSupportDetailStorageProvider;
            this.memoryCache = memoryCache;
        }

        /// <summary>
        /// Method will be invoked on each bot turn.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTurnAsync), turnContext);

                await base.OnTurnAsync(turnContext, cancellationToken);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error at {nameof(this.OnTurnAsync)}.");
                await base.OnTurnAsync(turnContext, cancellationToken);
                throw;
            }
        }

        /// <summary>
        /// Handle when a message is addressed to the bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// For more information on bot messaging in Teams, see the documentation
        /// https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-basics?tabs=dotnet#receive-a-message .
        /// </remarks>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                IMessageActivity activity = turnContext.Activity;
                this.RecordEvent(nameof(this.OnMessageActivityAsync), turnContext);
                await this.SendTypingIndicatorAsync(turnContext);

                switch (activity.Conversation.ConversationType)
                {
                    case PersonalConversationType:
                        await ActivityHelper.OnMessageActivityInPersonalChatAsync(activity, turnContext, this.logger, this.cardConfigurationStorageProvider, this.ticketGenerateStorageProvider, this.ticketDetailStorageProvider, this.microsoftAppCredentials, this.appBaseUrl, this.localizer, cancellationToken);
                        break;

                    case ChannelConversationType:
                        await ActivityHelper.OnMessageActivityInChannelAsync(message: activity, turnContext: turnContext, onCallSupportDetailSearchService: this.onCallSupportDetailSearchService, ticketDetailStorageProvider: this.ticketDetailStorageProvider, cardConfigurationStorageProvider: this.cardConfigurationStorageProvider, logger: this.logger, appBaseUrl: this.appBaseUrl, localizer: this.localizer, cancellationToken: cancellationToken);
                        break;

                    default:
                        throw new InvalidOperationException("Unexpected operation. Expected conversation type.");
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error processing message: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Invoke when a conversation update activity is received from the channel.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onconversationupdateactivityasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task OnConversationUpdateActivityAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnConversationUpdateActivityAsync), turnContext);
                IConversationUpdateActivity activity = turnContext.Activity;

                this.logger.LogInformation("Received conversationUpdate activity.");
                this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

                if (activity.MembersAdded?.Count == 0)
                {
                    this.logger.LogInformation("Ignoring conversationUpdate that was not a membersAdded event.");
                    return;
                }

                switch (activity.Conversation.ConversationType)
                {
                    case PersonalConversationType:
                        await ActivityHelper.OnMembersAddedToPersonalChatAsync(membersAdded: activity.MembersAdded, turnContext: turnContext, logger: this.logger, appBaseUrl: this.appBaseUrl, this.localizer);
                        return;

                    case ChannelConversationType:
                        await ActivityHelper.OnMembersAddedToTeamAsync(membersAdded: activity.MembersAdded, turnContext: turnContext, microsoftAppCredentials: this.microsoftAppCredentials, logger: this.logger, appBaseUrl: this.appBaseUrl, localizer: this.localizer, cancellationToken: cancellationToken);
                        return;

                    default:
                        throw new InvalidOperationException("Unexpected operation. Expected conversation type.");
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error processing conversationUpdate: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// When OnTurn method receives a fetch invoke activity on bot turn, it calls this method..
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            var activity = (Activity)turnContext.Activity;
            this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);

            try
            {
                var valuesforTaskModule = JsonConvert.DeserializeObject<AdaptiveCardAction>(((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase)?.ToString());
                string command = valuesforTaskModule.Command;

                switch (command.ToUpperInvariant())
                {
                    case Constants.EditRequestAction:
                        string ticketId = valuesforTaskModule.PostedValues;
                        var ticketDetail = await this.ticketDetailStorageProvider.GetTicketAsync(ticketId);
                        if (ticketDetail.TicketStatus == (int)TicketState.Closed)
                        {
                            return CardHelper.GetClosedErrorAdaptiveCard(this.localizer);
                        }

                        this.logger.LogInformation("Fetch and send edit request card.");
                        return CardHelper.GetEditTicketAdaptiveCard(cardConfigurationStorageProvider: this.cardConfigurationStorageProvider, ticketDetail: ticketDetail, localizer: this.localizer);

                    case Constants.ManageExpertsAction:
                        this.logger.LogInformation("Sending manage experts card.");
                        string customAPIAuthenticationToken = this.tokenHelper.GenerateAPIAuthToken(applicationBasePath: activity.ServiceUrl, fromId: activity.From.Id, jwtExpiryMinutes: 60);
                        return CardHelper.GetTaskModuleResponse(applicationBasePath: this.appBaseUrl, customAPIAuthenticationToken: customAPIAuthenticationToken, telemetryInstrumentationKey: this.telemetryClient.InstrumentationKey, activityId: valuesforTaskModule?.ActivityId, localizer: this.localizer);
                    default:
                        this.logger.LogInformation($"Invalid command for task module fetch activity.Command is : {command} ");
                        await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"));
                        return null;
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in fetching task module.");
                return null;
                throw;
            }
        }

        /// <summary>
        /// When OnTurn method receives a submit invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            var activity = (Activity)turnContext.Activity;
            this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);
            var valuesforTaskModule = JsonConvert.DeserializeObject<AdaptiveCardAction>(((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase)?.ToString());
            var editTicketDetail = JsonConvert.DeserializeObject<TicketDetail>(((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase)?.ToString());
            switch (valuesforTaskModule.Command.ToUpperInvariant())
            {
                case Constants.UpdateRequestAction:
                    var ticketDetail = await this.ticketDetailStorageProvider.GetTicketAsync(valuesforTaskModule.TicketId);
                    if (TicketHelper.ValidateRequestDetail(editTicketDetail, turnContext, ticketDetail))
                    {
                        ticketDetail.AdditionalProperties = CardHelper.ValidateAdditionalTicketDetails(((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase)?.ToString(), turnContext.Activity.Timestamp.Value.Offset);

                        // Update request card with user entered values.
                        ticketDetail = TicketHelper.GetUpdatedTicketDetails(turnContext, ticketDetail, editTicketDetail);
                        bool result = await this.ticketDetailStorageProvider.UpsertTicketAsync(ticketDetail);
                        if (!result)
                        {
                            this.logger.LogError("Error in storing new ticket details in table storage.");
                            await turnContext.SendActivityAsync(this.localizer.GetString("AzureStorageErrorText"));
                            return null;
                        }

                        // Send update audit trail message and request details card in personal chat and SME team.
                        this.logger.LogInformation($"Edited the ticket:{ticketDetail.TicketId}");
                        IMessageActivity smeEditNotification = MessageFactory.Text(string.Format(CultureInfo.InvariantCulture, this.localizer.GetString("SmeEditNotificationText"), ticketDetail.LastModifiedByName));

                        // Get card item element mappings
                        var cardElementMapping = await this.cardConfigurationStorageProvider.GetCardItemElementMappingAsync(ticketDetail.CardId);

                        IMessageActivity ticketDetailActivity = MessageFactory.Attachment(TicketCard.GetTicketDetailsForPersonalChatCard(cardElementMapping, ticketDetail, this.localizer, true));
                        ticketDetailActivity.Conversation = turnContext.Activity.Conversation;
                        ticketDetailActivity.Id = ticketDetail.RequesterTicketActivityId;
                        await turnContext.UpdateActivityAsync(ticketDetailActivity);
                        await CardHelper.UpdateSMECardAsync(turnContext, ticketDetail, smeEditNotification, this.appBaseUrl, cardElementMapping, this.localizer, this.logger, cancellationToken);
                    }
                    else
                    {
                        editTicketDetail.AdditionalProperties = CardHelper.ValidateAdditionalTicketDetails(additionalDetails: ((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase)?.ToString(), timeSpan: turnContext.Activity.Timestamp.Value.Offset);
                        return CardHelper.GetEditTicketAdaptiveCard(cardConfigurationStorageProvider: this.cardConfigurationStorageProvider, ticketDetail: editTicketDetail, localizer: this.localizer, existingTicketDetail: ticketDetail);
                    }

                    break;
                case Constants.UpdateExpertListAction:
                    var teamsChannelData = ((JObject)turnContext.Activity.ChannelData).ToObject<TeamsChannelData>();
                    var expertChannelId = teamsChannelData.Team == null ? this.teamId : teamsChannelData.Team.Id;
                    if (expertChannelId != this.teamId)
                    {
                        this.logger.LogInformation("Invalid team. Bot is not installed in this team.");
                        await turnContext.SendActivityAsync(this.localizer.GetString("InvalidTeamText"));
                        return null;
                    }

                    var onCallExpertsDetail = JsonConvert.DeserializeObject<OnCallExpertsDetail>(JObject.Parse(taskModuleRequest?.Data?.ToString())?.ToString());
                    await CardHelper.UpdateManageExpertsCardInTeamAsync(turnContext, onCallExpertsDetail, this.onCallSupportDetailSearchService, this.onCallSupportDetailStorageProvider, this.localizer);
                    await ActivityHelper.SendMentionActivityAsync(onCallExpertsDetail.OnCallExperts, turnContext: turnContext, logger: this.logger, localizer: this.localizer, memoryCache: this.memoryCache, cancellationToken: cancellationToken);
                    this.logger.LogInformation("Expert List has been updated");
                    return null;

                case Constants.CancelCommand:
                    return null;
            }

            return null;
        }

        /// <summary>
        /// Invoked when the user opens the messaging extension or searching any content in it.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Messaging extension response object to fill compose extension section.</returns>
        /// <remarks>
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionqueryasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionQuery query,
            CancellationToken cancellationToken)
        {
            IInvokeActivity turnContextActivity = turnContext?.Activity;
            try
            {
                MessagingExtensionQuery messageExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(turnContextActivity.Value.ToString());
                string searchQuery = SearchHelper.GetSearchQueryString(messageExtensionQuery);
                string onCallSMEUsers = string.Empty;
                turnContextActivity.TryGetChannelData<TeamsChannelData>(out var teamsChannelData);

                if (messageExtensionQuery.CommandId == Constants.ActiveCommandId || messageExtensionQuery.CommandId == Constants.ClosedCommandId)
                {
                    onCallSMEUsers = await CardHelper.GetOnCallSMEUserListAsync(turnContext, this.onCallSupportDetailSearchService, this.teamId, this.memoryCache, this.logger);
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = await SearchHelper.GetSearchResultAsync(searchQuery, messageExtensionQuery.CommandId, messageExtensionQuery.QueryOptions.Count, messageExtensionQuery.QueryOptions.Skip, this.ticketSearchService, this.localizer, turnContext.Activity.From.AadObjectId, onCallSMEUsers),
                    };
                }

                if (turnContext != null && teamsChannelData.Team != null && teamsChannelData.Team.Id == this.teamId && (messageExtensionQuery.CommandId == Constants.UrgentCommandId || messageExtensionQuery.CommandId == Constants.AssignedCommandId || messageExtensionQuery.CommandId == Constants.UnassignedCommandId))
                {
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = await SearchHelper.GetSearchResultAsync(searchQuery, messageExtensionQuery.CommandId, messageExtensionQuery.QueryOptions.Count, messageExtensionQuery.QueryOptions.Skip, this.ticketSearchService, this.localizer),
                    };
                }

                return new MessagingExtensionResponse
                {
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Text = this.localizer.GetString("InvalidTeamText"),
                        Type = "message",
                    },
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to handle the messaging extension command {turnContextActivity.Name}: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Records event occurred in the application in Application Insights telemetry client.
        /// </summary>
        /// <param name="eventName"> Name of the event.</param>
        /// <param name="turnContext"> Context object containing information cached for a single turn of conversation with a user.</param>
        private void RecordEvent(string eventName, ITurnContext turnContext)
        {
            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", turnContext.Activity.From.AadObjectId },
                { "tenantId", turnContext.Activity.Conversation.TenantId },
            });
        }

        /// <summary>
        /// Send typing indicator to the user.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>A task that represents typing indicator activity.</returns>
        private async Task SendTypingIndicatorAsync(ITurnContext turnContext)
        {
            try
            {
                var typingActivity = turnContext.Activity.CreateReply();
                typingActivity.Type = ActivityTypes.Typing;
                await turnContext.SendActivityAsync(typingActivity);
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                // Do not fail on errors sending the typing indicator
                this.logger.LogWarning(ex, "Failed to send a typing indicator.");
            }
        }
    }
}