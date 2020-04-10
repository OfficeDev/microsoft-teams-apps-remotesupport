// <copyright file="ActivityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Xml;
    using AdaptiveCards;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RemoteSupport;
    using Microsoft.Teams.Apps.RemoteSupport.Cards;
    using Microsoft.Teams.Apps.RemoteSupport.Common;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Providers;
    using Microsoft.Teams.Apps.RemoteSupport.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Class that handles Bot activity helper methods.
    /// </summary>
    public static class ActivityHelper
    {
        /// <summary>
        /// RequestType - text that triggers severity action by SME.
        /// </summary>
        private const string RequestTypeText = "RequestType";

        /// <summary>
        /// Handle message activity in channel.
        /// </summary>
        /// <param name="message">A message in a conversation.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="onCallSupportDetailSearchService">Provider to search on call support details in Azure Table Storage.</param>
        /// <param name="ticketDetailStorageProvider">Provider to store ticket details to Azure Table Storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="appBaseUrl">Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        internal static async Task OnMessageActivityInChannelAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            IOnCallSupportDetailSearchService onCallSupportDetailSearchService,
            ITicketDetailStorageProvider ticketDetailStorageProvider,
            ILogger<RemoteSupportActivityHandler> logger,
            string appBaseUrl,
            IStringLocalizer<Strings> localizer,
            CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            if (!string.IsNullOrEmpty(message.ReplyToId) && message.Value != null && ((JObject)message.Value).HasValues)
            {
                logger.LogInformation($"Card submit in channel {message.Value?.ToString()}");
                await OnAdaptiveCardSubmitInChannelAsync(message: message, turnContext: turnContext, ticketDetailStorageProvider: ticketDetailStorageProvider, logger: logger, appBaseUrl: appBaseUrl, localizer: localizer, cancellationToken: cancellationToken);
                return;
            }

            turnContext.Activity.RemoveRecipientMention();
            string text = turnContext.Activity.Text.Trim();

            switch (text.ToUpperInvariant())
            {
                case Constants.ManageExpertsAction:
                    // Get on call support data from storage
                    var onCallSupportDetails = await onCallSupportDetailSearchService.SearchOnCallSupportTeamAsync(searchQuery: string.Empty, count: 10);
                    var onCallSMEDetailActivity = MessageFactory.Attachment(OnCallSMEDetailCard.GetOnCallSMEDetailCard(onCallSupportDetails, localizer));
                    var result = await turnContext.SendActivityAsync(onCallSMEDetailActivity);

                    // Add activityId in the data which will be posted to task module in future after clicking on Manage button.
                    AdaptiveCard adaptiveCard = (AdaptiveCard)onCallSMEDetailActivity.Attachments?[0].Content;
                    AdaptiveCardAction cardAction = (AdaptiveCardAction)((AdaptiveSubmitAction)adaptiveCard?.Actions?[0]).Data;
                    cardAction.ActivityId = result.Id;

                    // Refresh manage experts card with activity Id bound to manage button.
                    onCallSMEDetailActivity.Id = result.Id;
                    onCallSMEDetailActivity.ReplyToId = result.Id;
                    await turnContext.UpdateActivityAsync(onCallSMEDetailActivity);

                    break;

                default:
                    logger.LogInformation("Unrecognized input in channel.");
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeTeamCard.GetCard(appBaseUrl, localizer)));
                    break;
            }
        }

        /// <summary>
        /// Handle adaptive card submit in channel.
        /// Updates the ticket status based on the user submission.
        /// </summary>
        /// <param name="message">A message in a conversation.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="ticketDetailStorageProvider">Provider to store ticket details to Azure Table Storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="appBaseUrl">Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        internal static async Task OnAdaptiveCardSubmitInChannelAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            ITicketDetailStorageProvider ticketDetailStorageProvider,
            ILogger<RemoteSupportActivityHandler> logger,
            string appBaseUrl,
            IStringLocalizer<Strings> localizer,
            CancellationToken cancellationToken)
        {
            string smeNotification;
            IMessageActivity userNotification;
            ChangeTicketStatus payload = ((JObject)message.Value).ToObject<ChangeTicketStatus>();
            payload.Action = payload.RequestType == null ? payload.Action : RequestTypeText;
            logger.LogInformation($"Received submit:  action={payload.Action} ticketId={payload.TicketId}");

            // Get the ticket from the data store.
            TicketDetail ticketData = await ticketDetailStorageProvider.GetTicketAsync(payload.TicketId);
            if (ticketData == null)
            {
                await turnContext.SendActivityAsync($"Ticket {payload.TicketId} was not found in the data store");
                logger.LogInformation($"Ticket {payload.TicketId} was not found in the data store");
                return;
            }

            // Update the ticket based on the payload.
            switch (payload.Action)
            {
                case ChangeTicketStatus.ReopenAction:
                    ticketData.TicketStatus = (int)TicketState.Unassigned;
                    ticketData.AssignedToName = null;
                    ticketData.AssignedToObjectId = null;
                    ticketData.ClosedOn = null;
                    smeNotification = localizer.GetString("SmeUnassignedStatus", message.From.Name);
                    userNotification = MessageFactory.Text(localizer.GetString("ReopenedTicketUserNotification", ticketData.TicketId));
                    break;

                case ChangeTicketStatus.CloseAction:
                    ticketData.TicketStatus = (int)TicketState.Closed;
                    ticketData.ClosedByName = message.From.Name;
                    ticketData.ClosedOn = message.From.AadObjectId;
                    smeNotification = localizer.GetString("SmeClosedStatus", message.From.Name);
                    userNotification = MessageFactory.Text(localizer.GetString("ClosedTicketUserNotification", ticketData.TicketId));
                    break;

                case ChangeTicketStatus.AssignToSelfAction:
                    ticketData.TicketStatus = (int)TicketState.Assigned;
                    ticketData.AssignedToName = message.From.Name;
                    ticketData.AssignedToObjectId = message.From.AadObjectId;
                    ticketData.ClosedOn = null;
                    smeNotification = localizer.GetString("SmeAssignedStatus", message.From.Name);
                    userNotification = MessageFactory.Text(localizer.GetString("AssignedTicketUserNotification", ticketData.TicketId));
                    break;

                case ChangeTicketStatus.RequestTypeAction:
                    ticketData.Severity = (int)(TicketSeverity)Enum.Parse(typeof(TicketSeverity), payload.RequestType);
                    ticketData.RequestType = payload.RequestType;
                    logger.LogInformation($"Received submit:  action={payload.RequestType} ticketId={payload.TicketId}");
                    smeNotification = localizer.GetString("SmeSeveritystatus", ticketData.RequestType, message.From.Name);
                    userNotification = MessageFactory.Text(localizer.GetString("RequestActionTicketUserNotification", ticketData.TicketId));
                    break;

                default:
                    logger.LogInformation($"Unknown status command {payload.Action}", SeverityLevel.Warning);
                    return;
            }

            ticketData.LastModifiedByName = message.From.Name;
            ticketData.LastModifiedByObjectId = message.From.AadObjectId;
            ticketData.LastModifiedOn = DateTime.UtcNow;

            await ticketDetailStorageProvider.UpsertTicketAsync(ticketData);
            logger.LogInformation($"Ticket {ticketData.TicketId} updated to status ({ticketData.TicketStatus}, {ticketData.AssignedToObjectId}) in store");

            // Update the card in the SME team.
            Activity updateCardActivity = new Activity(ActivityTypes.Message)
            {
                Id = ticketData.SmeTicketActivityId,
                Conversation = new ConversationAccount { Id = ticketData.SmeConversationId },
                Attachments = new List<Attachment> { new SmeTicketCard(ticketData).GetTicketDetailsForSMEChatCard(ticketData, appBaseUrl, localizer) },
            };
            ResourceResponse updateResponse = await turnContext.UpdateActivityAsync(updateCardActivity, cancellationToken);
            logger.LogInformation($"Card for ticket {ticketData.TicketId} updated to status ({ticketData.TicketStatus}, {ticketData.AssignedToObjectId}), activityId = {updateResponse.Id}");

            // Post update to user and SME team thread.
            if (!string.IsNullOrEmpty(smeNotification))
            {
                ResourceResponse smeResponse = await turnContext.SendActivityAsync(smeNotification);
                logger.LogInformation($"SME team notified of update to ticket {ticketData.TicketId}, activityId = {smeResponse.Id}");
            }

            if (userNotification != null)
            {
                userNotification.Conversation = new ConversationAccount { Id = ticketData.RequesterConversationId };
                ResourceResponse[] userResponse = await turnContext.Adapter.SendActivitiesAsync(turnContext, new Activity[] { (Activity)userNotification }, cancellationToken);
                logger.LogInformation($"User notified of update to ticket {ticketData.TicketId}, activityId = {userResponse.FirstOrDefault()?.Id}");
            }
        }

        /// <summary>
        /// Handle when a message is addressed to the bot in personal scope.
        /// </summary>
        /// <param name="message">Message activity of bot.</param>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client. </param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="cardConfigurationStorageProvider">Provider to search card configuration details in Azure Table Storage.</param>
        /// <param name="environment">Hosting environment.</param>
        /// <param name="ticketGenerateStorageProvider">Provider to get ticket id to Azure Table Storage.</param>
        /// <param name="ticketDetailStorageProvider">Provider to store ticket details to Azure Table Storage.</param>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="appBaseUrl">Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns> A task that represents the work queued to execute for user message activity to bot.</returns>
        internal static async Task OnMessageActivityInPersonalChatAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            TelemetryClient telemetryClient,
            ILogger<RemoteSupportActivityHandler> logger,
            ICardConfigurationStorageProvider cardConfigurationStorageProvider,
            IHostingEnvironment environment,
            ITicketIdGeneratorStorageProvider ticketGenerateStorageProvider,
            ITicketDetailStorageProvider ticketDetailStorageProvider,
            MicrosoftAppCredentials microsoftAppCredentials,
            string appBaseUrl,
            IStringLocalizer<Strings> localizer,
            CancellationToken cancellationToken)
        {
            if (!string.IsNullOrEmpty(message.ReplyToId) && message.Value != null && ((JObject)message.Value).HasValues)
            {
                telemetryClient.TrackTrace("Card submitted in 1:1 chat.");
                await OnAdaptiveCardSubmitInPersonalChatAsync(message: message, turnContext: turnContext, ticketGenerateStorageProvider: ticketGenerateStorageProvider, ticketDetailStorageProvider: ticketDetailStorageProvider, cardConfigurationStorageProvider: cardConfigurationStorageProvider, microsoftAppCredentials: microsoftAppCredentials, logger: logger, appBaseUrl: appBaseUrl, environment: environment, localizer: localizer, cancellationToken: cancellationToken);
                return;
            }

            string text = (turnContext.Activity.Text ?? string.Empty).Trim().ToUpperInvariant();
            switch (text)
            {
                case Constants.NewRequestAction:
                    logger.LogInformation("New request action called.");
                    CardConfigurationEntity cardTemplateJson = await cardConfigurationStorageProvider.GetConfigurationAsync();
                    IMessageActivity newTicketActivity = MessageFactory.Attachment(TicketCard.GetNewTicketCard(environment, cardTemplateJson, localizer));
                    await turnContext.SendActivityAsync(newTicketActivity);
                    break;

                case Constants.NoCommand:
                    return;

                default:
                    if (turnContext.Activity.Attachments == null || turnContext.Activity.Attachments.Count == 0)
                    {
                        // In case of ME when user clicks on closed or active requests the bot posts adaptive card of request details we don't have to consider this as invalid command.
                        logger.LogInformation("Unrecognized input in End User.");
                        await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeCard.GetCard(appBaseUrl, localizer)));
                    }

                    break;
            }
        }

        /// <summary>
        /// Method Handle adaptive card submit in 1:1 chat and Send new ticket details to SME team.
        /// </summary>
        /// <param name="message">Message activity of bot.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="ticketGenerateStorageProvider">Provider to get ticket id to Azure Table Storage.</param>
        /// <param name="ticketDetailStorageProvider">Provider to store ticket details to Azure Table Storage.</param>
        /// <param name="cardConfigurationStorageProvider">Provider to search card configuration details in Azure Table Storage.</param>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="appBaseUrl">Represents the Application base Uri.</param>
        /// <param name="environment">Hosting environment.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that handles submit action in 1:1 chat.</returns>
        internal static async Task OnAdaptiveCardSubmitInPersonalChatAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            ITicketIdGeneratorStorageProvider ticketGenerateStorageProvider,
            ITicketDetailStorageProvider ticketDetailStorageProvider,
            ICardConfigurationStorageProvider cardConfigurationStorageProvider,
            MicrosoftAppCredentials microsoftAppCredentials,
            ILogger<RemoteSupportActivityHandler> logger,
            string appBaseUrl,
            IHostingEnvironment environment,
            IStringLocalizer<Strings> localizer,
            CancellationToken cancellationToken)
        {
            IMessageActivity endUserUpdateCard;

            switch (message.Text.ToUpperInvariant())
            {
                case Constants.SendRequestAction:
                    TicketDetail newTicketDetail = JsonConvert.DeserializeObject<TicketDetail>(message.Value?.ToString());
                    if (TicketHelper.ValidateRequestDetail(newTicketDetail, turnContext))
                    {
                        AdaptiveCardAction cardDetail = ((JObject)message.Value).ToObject<AdaptiveCardAction>();
                        logger.LogInformation("Adding new request with additional details.");
                        var ticketTd = await ticketGenerateStorageProvider.GetTicketIdAsync();

                        // Update new request with additional details.
                        var userDetails = await GetUserDetailsInPersonalChatAsync(turnContext, cancellationToken);
                        newTicketDetail.TicketId = ticketTd.ToString(CultureInfo.InvariantCulture);
                        newTicketDetail = TicketHelper.GetNewTicketDetails(turnContext: turnContext, ticketDetail: newTicketDetail, ticketAdditionalDetails: message.Value?.ToString(), cardId: cardDetail.CardId, member: userDetails);
                        bool result = await ticketDetailStorageProvider.UpsertTicketAsync(newTicketDetail);
                        if (!result)
                        {
                            logger.LogError("Error in storing new ticket details in table storage.");
                            await turnContext.SendActivityAsync(localizer.GetString("AzureStorageErrorText"));
                            return;
                        }

                        logger.LogInformation("New request created with ticket Id:" + newTicketDetail.TicketId);
                        endUserUpdateCard = MessageFactory.Attachment(TicketCard.GetTicketDetailsForPersonalChatCard(newTicketDetail, localizer, false));
                        await CardHelper.SendRequestCardToSMEChannelAsync(turnContext: turnContext, ticketDetail: newTicketDetail, logger: logger, ticketDetailStorageProvider: ticketDetailStorageProvider, applicationBasePath: appBaseUrl, localizer, teamId: cardDetail?.TeamId, microsoftAppCredentials: microsoftAppCredentials, cancellationToken: cancellationToken);
                        await CardHelper.UpdateRequestCardForEndUserAsync(turnContext, endUserUpdateCard);

                        await turnContext.SendActivityAsync(MessageFactory.Text(localizer.GetString("EndUserNotificationText", newTicketDetail.TicketId)));
                    }
                    else
                    {
                        // Update card with validation message.
                        var additionalProperties = message.Value?.ToString();
                        CardConfigurationEntity cardTemplateJson = await cardConfigurationStorageProvider.GetConfigurationAsync();
                        endUserUpdateCard = MessageFactory.Attachment(TicketCard.GetNewTicketCard(environment: environment, cardConfiuration: cardTemplateJson, localizer: localizer, showValidationMessage: true, ticketDetail: newTicketDetail, ticketAdditionalDetails: additionalProperties));
                        await CardHelper.UpdateRequestCardForEndUserAsync(turnContext, endUserUpdateCard);
                    }

                    break;

                case Constants.WithdrawRequestAction:
                    var payload = ((JObject)message.Value).ToObject<AdaptiveCardAction>();
                    endUserUpdateCard = MessageFactory.Attachment(WithdrawCard.GetCard(payload.PostedValues, localizer));

                    // Get the ticket from the data store.
                    TicketDetail ticketDetail = await ticketDetailStorageProvider.GetTicketAsync(payload.PostedValues);
                    if (ticketDetail.TicketStatus == (int)TicketState.Closed)
                    {
                        await turnContext.SendActivityAsync(localizer.GetString("WithdrawErrorMessage"));
                        return;
                    }

                    ticketDetail.LastModifiedByName = message.From.Name;
                    ticketDetail.LastModifiedByObjectId = message.From.AadObjectId;
                    ticketDetail.TicketStatus = (int)TicketState.Withdrawn;
                    bool success = await ticketDetailStorageProvider.UpsertTicketAsync(ticketDetail);
                    if (!success)
                    {
                        logger.LogError("Error in updating ticket details in table storage.");
                        await turnContext.SendActivityAsync(localizer.GetString("AzureStorageErrorText"));
                        return;
                    }

                    logger.LogInformation("Withdrawn the ticket:" + ticketDetail.TicketId);
                    IMessageActivity smeWithdrawNotification = MessageFactory.Text(localizer.GetString("SmeWithdrawNotificationText", ticketDetail.RequesterName));
                    await CardHelper.UpdateSMECardAsync(turnContext, ticketDetail, smeWithdrawNotification, appBaseUrl, localizer, logger, cancellationToken);
                    await CardHelper.UpdateRequestCardForEndUserAsync(turnContext, endUserUpdateCard);
                    break;
            }
        }

        /// <summary>
        /// Handle members added conversationUpdate event in team.
        /// </summary>
        /// <param name="membersAdded">Channel account information needed to route a message.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="appBaseUrl">Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        internal static async Task OnMembersAddedToTeamAsync(
           IList<ChannelAccount> membersAdded,
           ITurnContext<IConversationUpdateActivity> turnContext,
           MicrosoftAppCredentials microsoftAppCredentials,
           ILogger<RemoteSupportActivityHandler> logger,
           string appBaseUrl,
           IStringLocalizer<Strings> localizer,
           CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            var activity = turnContext.Activity;
            if (membersAdded.Any(channelAccount => channelAccount.Id == activity.Recipient.Id))
            {
                // Bot was added to a team
                logger.LogInformation($"Bot added to team {activity.Conversation.Id}");
                var teamDetails = ((JObject)turnContext.Activity.ChannelData).ToObject<TeamsChannelData>();
                var teamWelcomeCardAttachment = WelcomeTeamCard.GetCard(appBaseUrl, localizer);
                await CardHelper.SendCardToTeamAsync(turnContext, teamWelcomeCardAttachment, teamDetails.Team.Id, microsoftAppCredentials, cancellationToken);
            }
        }

        /// <summary>
        /// Handle 1:1 chat with members who started chat for the first time.
        /// </summary>
        /// <param name="membersAdded">Channel account information needed to route a message.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="appBaseUrl">Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        internal static async Task OnMembersAddedToPersonalChatAsync(
            IList<ChannelAccount> membersAdded,
            ITurnContext<IConversationUpdateActivity> turnContext,
            ILogger<RemoteSupportActivityHandler> logger,
            string appBaseUrl,
            IStringLocalizer<Strings> localizer)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            var activity = turnContext.Activity;
            if (membersAdded.Any(channelAccount => channelAccount.Id == activity.Recipient.Id))
            {
                // User started chat with the bot in personal scope, for the first time.
                logger.LogInformation($"Bot added to 1:1 chat {activity.Conversation.Id}");
                await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeCard.GetCard(appBaseUrl, localizer)));
            }
        }

        /// <summary>
        /// Method mentions user in respective channel of which they are part after modifying experts list.
        /// </summary>
        /// <param name="onCallExpertsEmails">Collection of on call expert emails.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that sends notification in newly created channel and mention its members.</returns>
        internal static async Task<Activity> SendMentionActivityAsync(
            List<string> onCallExpertsEmails,
            ITurnContext<IInvokeActivity> turnContext,
            ILogger<RemoteSupportActivityHandler> logger,
            IStringLocalizer<Strings> localizer,
            CancellationToken cancellationToken)
        {
            try
            {
                var mentionText = new StringBuilder();
                var entities = new List<Entity>();
                var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
                var channelMembers = await TeamsInfo.GetTeamMembersAsync(turnContext, teamsDetails.Id, cancellationToken);

                var onCallExpertDetails = channelMembers.Where(member => onCallExpertsEmails.Contains(member.Email)).Select(member => new ChannelAccount { Id = member.Id, Name = member.Name });

                foreach (var onCallExpert in onCallExpertDetails)
                {
                    var mention = new Mention
                    {
                        Mentioned = new ChannelAccount()
                        {
                            Id = onCallExpert.Id,
                            Name = onCallExpert.Name,
                        },
                        Text = $"<at>{XmlConvert.EncodeName(onCallExpert.Name)}</at>",
                    };
                    entities.Add(mention);
                    mentionText = string.IsNullOrEmpty(mentionText.ToString()) ? mentionText.Append(mention.Text) : mentionText.Append(", ").Append(mention.Text);
                }

                logger.LogInformation("Send message with names mentioned in team channel.");
                var replyActivity = string.IsNullOrEmpty(mentionText.ToString()) ? MessageFactory.Text(localizer.GetString("OnCallListUpdateMessage")) : MessageFactory.Text(localizer.GetString("OnCallExpertMentionText", mentionText.ToString()));
                replyActivity.Entities = entities;
                await turnContext.SendActivityAsync(replyActivity, cancellationToken);
                return null;
            }
            #pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
            #pragma warning restore CA1031 // Do not catch general exception types
            {
                logger.LogError(ex, $"Error while mentioning channel member in respective channels.");
                return null;
            }
        }

        /// <summary>
        /// Get the account details of the user in a 1:1 chat with the bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        internal static async Task<TeamsChannelAccount> GetUserDetailsInPersonalChatAsync(
          ITurnContext<IMessageActivity> turnContext,
          CancellationToken cancellationToken)
        {
            var members = await ((BotFrameworkAdapter)turnContext.Adapter).GetConversationMembersAsync(turnContext, cancellationToken);
            return JsonConvert.DeserializeObject<TeamsChannelAccount>(JsonConvert.SerializeObject(members[0]));
        }

        /// <summary>
        /// Verify if the tenant Id in the message is the same tenant Id used when application was configured.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="tenantId">Represents unique id of a Tenant.</param>
        /// <returns>True if context is from expected tenant else false.</returns>
        internal static bool IsActivityFromExpectedTenant(ITurnContext turnContext, string tenantId)
        {
            return turnContext.Activity.Conversation.TenantId == tenantId;
        }
    }
}
