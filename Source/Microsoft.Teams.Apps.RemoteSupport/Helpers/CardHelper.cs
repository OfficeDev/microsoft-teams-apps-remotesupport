// <copyright file="CardHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RemoteSupport.Cards;
    using Microsoft.Teams.Apps.RemoteSupport.Common;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Providers;
    using Microsoft.Teams.Apps.RemoteSupport.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Class that handles the card configuration.
    /// </summary>
    public static class CardHelper
    {
        /// <summary>
        /// Task module height.
        /// </summary>
        private const int TaskModuleHeight = 460;

        /// <summary>
        /// Represents the task module width.
        /// </summary>
        private const int TaskModuleWidth = 600;

        /// <summary>
        /// Task module height.
        /// </summary>
        private const int ErrorMessageTaskModuleHeight = 100;

        /// <summary>
        /// Represents the task module width.
        /// </summary>
        private const int ErrorMessageTaskModuleWidth = 400;

        /// <summary>
        /// Update request card in end user conversation.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="endUserUpdateCard"> End user request details card which is to be updated in end user conversation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public static async Task<bool> UpdateRequestCardForEndUserAsync(ITurnContext turnContext, IMessageActivity endUserUpdateCard)
        {
            if (endUserUpdateCard != null)
            {
                endUserUpdateCard.Id = turnContext?.Activity.ReplyToId;
                endUserUpdateCard.Conversation = turnContext.Activity.Conversation;
                await turnContext.UpdateActivityAsync(endUserUpdateCard);
                return true;
            }
            else
            {
                throw new Exception("Error while updating card in end user conversation.");
            }
        }

        /// <summary>
        /// Get task module response.
        /// </summary>
        /// <param name="applicationBasePath">Represents the Application base Uri.</param>
        /// <param name="customAPIAuthenticationToken">JWT token.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="activityId">Task module activity Id.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns task module response.</returns>
        public static TaskModuleResponse GetTaskModuleResponse(string applicationBasePath, string customAPIAuthenticationToken, TelemetryClient telemetryClient, string activityId, IStringLocalizer<Strings> localizer)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Url = $"{applicationBasePath}/manage-experts?token={customAPIAuthenticationToken}&telemetry={telemetryClient?.InstrumentationKey}&activityId={activityId}&theme=" + "{theme}&locale=" + "{locale}",
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = localizer.GetString("ManageExpertsTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Gets edit ticket details adaptive card.
        /// </summary>
        /// <param name="cardConfigurationStorageProvider">Card configuration.</param>
        /// <param name="ticketDetail">Details of the ticket to be edited.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="existingTicketDetail">Existing ticket details.</param>
        /// <returns>Returns edit ticket adaptive card.</returns>
        public static TaskModuleResponse GetEditTicketAdaptiveCard(ICardConfigurationStorageProvider cardConfigurationStorageProvider, TicketDetail ticketDetail, IStringLocalizer<Strings> localizer, TicketDetail existingTicketDetail = null)
        {
            var cardTemplate = cardConfigurationStorageProvider?.GetConfigurationsByCardIdAsync(ticketDetail?.CardId).Result;
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = EditRequestCard.GetEditRequestCard(ticketDetail, cardTemplate, localizer, existingTicketDetail),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = localizer.GetString("EditRequestTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Gets error message details adaptive card.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns edit ticket adaptive card.</returns>
        public static TaskModuleResponse GetClosedErrorAdaptiveCard(IStringLocalizer<Strings> localizer)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = EditRequestCard.GetClosedErrorCard(localizer),
                        Height = ErrorMessageTaskModuleHeight,
                        Width = ErrorMessageTaskModuleWidth,
                        Title = localizer.GetString("EditRequestTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Send card to SME channel and storage conversation details in storage.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="ticketDetail">Ticket details entered by user.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="ticketDetailStorageProvider">Provider to store ticket details to Azure Table Storage.</param>
        /// <param name="applicationBasePath">Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamId">Represents unique id of a Team.</param>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Returns message in a conversation.</returns>
        public static async Task<ConversationResourceResponse> SendRequestCardToSMEChannelAsync(
            ITurnContext<IMessageActivity> turnContext,
            TicketDetail ticketDetail,
            ILogger<RemoteSupportActivityHandler> logger,
            ITicketDetailStorageProvider ticketDetailStorageProvider,
            string applicationBasePath,
            IStringLocalizer<Strings> localizer,
            string teamId,
            MicrosoftAppCredentials microsoftAppCredentials,
            CancellationToken cancellationToken)
        {
            Attachment smeTeamCard = new SmeTicketCard(ticketDetail).GetTicketDetailsForSMEChatCard(ticketDetail, applicationBasePath, localizer);
            ConversationResourceResponse resourceResponse = await SendCardToTeamAsync(turnContext, smeTeamCard, teamId, microsoftAppCredentials, cancellationToken);

            if (resourceResponse == null)
            {
                logger.LogError("Error while sending card to team.");
                return null;
            }

            // Update SME team conversation details in storage.
            ticketDetail.SmeTicketActivityId = resourceResponse.ActivityId;
            ticketDetail.SmeConversationId = resourceResponse.Id;
            bool result = await ticketDetailStorageProvider?.UpsertTicketAsync(ticketDetail);

            if (!result)
            {
                logger.LogError("Error while saving SME conversation details in storage.");
            }

            return resourceResponse;
        }

        /// <summary>
        /// Send the given attachment to the specified team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cardToSend">The card to send.</param>
        /// <param name="teamId">Team id to which the message is being sent.</param>
        /// <param name="microsoftAppCredentials">Microsoft Application credentials for Bot/ME.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns><see cref="Task"/>That resolves to a <see cref="ConversationResourceResponse"/>Send a attachment.</returns>
        public static async Task<ConversationResourceResponse> SendCardToTeamAsync(
            ITurnContext turnContext,
            Attachment cardToSend,
            string teamId,
            MicrosoftAppCredentials microsoftAppCredentials,
            CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            ConversationParameters conversationParameters = new ConversationParameters
            {
                Activity = (Activity)MessageFactory.Attachment(cardToSend),
                ChannelData = new TeamsChannelData { Channel = new ChannelInfo(teamId) },
            };

            TaskCompletionSource<ConversationResourceResponse> taskCompletionSource = new TaskCompletionSource<ConversationResourceResponse>();
            await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                null, // If we set channel = "msteams", there is an error as preinstalled middle-ware expects ChannelData to be present.
                turnContext.Activity.ServiceUrl,
                microsoftAppCredentials,
                conversationParameters,
                (newTurnContext, newCancellationToken) =>
                {
                    Activity activity = newTurnContext.Activity;
                    taskCompletionSource.SetResult(new ConversationResourceResponse
                    {
                        Id = activity.Conversation.Id,
                        ActivityId = activity.Id,
                        ServiceUrl = activity.ServiceUrl,
                    });
                    return Task.CompletedTask;
                },
                cancellationToken);

            return await taskCompletionSource.Task;
        }

        /// <summary>
        /// Gets the email id's of the SME uses who are available for oncallSupport.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="onCallSupportDetailSearchService">Provider to search on call support details in Azure Table Storage.</param>
        /// <param name="teamId">Team id to which the message is being sent.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <returns>string with appended email id's.</returns>
        public static async Task<string> GetOnCallSMEuserListAsync(ITurnContext<IInvokeActivity> turnContext, IOnCallSupportDetailSearchService onCallSupportDetailSearchService, string teamId, ILogger<RemoteSupportActivityHandler> logger)
        {
            try
            {
                var teamsChannelAccounts = await TeamsInfo.GetTeamMembersAsync(turnContext, teamId, CancellationToken.None);
                var onCallSupportDetails = await onCallSupportDetailSearchService?.SearchOnCallSupportTeamAsync(string.Empty, 1);
                string onCallSMEUsers = string.Empty;
                if (onCallSupportDetails != null && onCallSupportDetails.Any())
                {
                    var onCallSMEDetail = JsonConvert.DeserializeObject<List<OnCallSMEDetail>>(onCallSupportDetails.First().OnCallSMEs);
                    if (onCallSMEDetail != null)
                    {
                        foreach (var onCallSME in onCallSMEDetail)
                        {
                            onCallSMEUsers += string.IsNullOrEmpty(onCallSMEUsers) ? teamsChannelAccounts.FirstOrDefault(teamsChannelAccount => teamsChannelAccount.AadObjectId == onCallSME.ObjectId)?.Email : "," + teamsChannelAccounts.FirstOrDefault(teamsChannelAccount => teamsChannelAccount.AadObjectId == onCallSME.ObjectId)?.Email;
                        }
                    }
                }

                return onCallSMEUsers;
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception ex)
#pragma warning restore CA1031 // Do not catch general exception types
            {
                logger.LogError(ex, "Error in getting the oncallSMEUsers list.");
            }

            return null;
        }

        /// <summary>
        /// Method updates experts card in team after modifying on call experts list.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="onCallExpertsDetail">Details of on call support experts updated.</param>
        /// <param name="onCallSupportDetailSearchService">Provider to search on call support details in Azure Table Storage.</param>
        /// <param name="onCallSupportDetailStorageProvider"> Provider for fetching and storing information about on call support in storage table.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>A task that sends notification in newly created channel and mention its members.</returns>
        public static async Task UpdateManageExpertsCardInTeamAsync(ITurnContext<IInvokeActivity> turnContext, OnCallExpertsDetail onCallExpertsDetail, IOnCallSupportDetailSearchService onCallSupportDetailSearchService, IOnCallSupportDetailStorageProvider onCallSupportDetailStorageProvider, IStringLocalizer<Strings> localizer)
        {
            // Get last 10 updated on call support data from storage.
            // This is required because search service refresh interval is 10 minutes. So we need to get latest entry stored in storage from storage provider and append previous 9 updated records to it in order to show on screen.
            var previousOnCallSupportDetails = await onCallSupportDetailSearchService?.SearchOnCallSupportTeamAsync(string.Empty, 9);
            var currentOnCallSupportDetails = await onCallSupportDetailStorageProvider?.GetOnCallSupportDetailAsync(onCallExpertsDetail?.OnCallSupportId);

            List<OnCallSupportDetail> onCallSupportDetails = new List<OnCallSupportDetail>
            {
                currentOnCallSupportDetails,
            };
            onCallSupportDetails.AddRange(previousOnCallSupportDetails);

            // Replace message id in conversation id with card activity id to be refreshed.
            var conversationId = turnContext?.Activity.Conversation.Id;
            conversationId = conversationId?.Replace(turnContext.Activity.Conversation.Id.Split(';')[1].Split("=")[1], onCallExpertsDetail?.OnCallSupportCardActivityId, StringComparison.OrdinalIgnoreCase);
            var onCallSMEDetailCardAttachment = OnCallSMEDetailCard.GetOnCallSMEDetailCard(onCallSupportDetails, localizer);

            // Add activityId in the data which will be posted to task module in future after clicking on Manage button.
            AdaptiveCard adaptiveCard = (AdaptiveCard)onCallSMEDetailCardAttachment.Content;
            AdaptiveCardAction cardAction = (AdaptiveCardAction)((AdaptiveSubmitAction)adaptiveCard?.Actions?[0]).Data;
            cardAction.ActivityId = onCallExpertsDetail?.OnCallSupportCardActivityId;

            // Update the card in the SME team with updated on call experts list.
            var updateExpertsCardActivity = new Activity(ActivityTypes.Message)
            {
                Id = onCallExpertsDetail?.OnCallSupportCardActivityId,
                ReplyToId = onCallExpertsDetail?.OnCallSupportCardActivityId,
                Conversation = new ConversationAccount { Id = conversationId },
                Attachments = new List<Attachment> { onCallSMEDetailCardAttachment },
            };
            await turnContext.UpdateActivityAsync(updateExpertsCardActivity);
        }

        /// <summary>
        /// Method to update the SME Card and gives corresponding notification.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="ticketDetail"> Ticket details entered by user.</param>
        /// <param name="messageActivity">Message activity of bot.</param>
        /// <param name="applicationBasePath"> Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="logger">Telemetry logger.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>task that updates card.</returns>
        public static async Task UpdateSMECardAsync(ITurnContext turnContext, TicketDetail ticketDetail, IMessageActivity messageActivity, string applicationBasePath, IStringLocalizer<Strings> localizer, ILogger<RemoteSupportActivityHandler> logger, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                messageActivity = messageActivity ?? throw new ArgumentNullException(nameof(messageActivity));

                // Update the card in the SME team.
                var updateCardActivity = new Activity(ActivityTypes.Message)
                {
                    Id = ticketDetail?.SmeTicketActivityId,
                    Conversation = new ConversationAccount { Id = ticketDetail.SmeConversationId },
                    Attachments = new List<Attachment> { new SmeTicketCard(ticketDetail).GetTicketDetailsForSMEChatCard(ticketDetail, applicationBasePath, localizer) },
                };
                await turnContext.Adapter.UpdateActivityAsync(turnContext, updateCardActivity, cancellationToken);
                messageActivity.Conversation = new ConversationAccount { Id = ticketDetail.SmeConversationId };
                await turnContext.Adapter.SendActivitiesAsync(turnContext, new Activity[] { (Activity)messageActivity }, cancellationToken);
            }
            catch (ErrorResponseException ex)
            {
                if (ex.Body.Error.Code == "ConversationNotFound")
                {
                    // exception could also be thrown by bot adapter if updated activity is same as current
                    logger.LogError(ex, $"failed to update activity due to conversation id not found {nameof(UpdateSMECardAsync)}");
                }

                logger.LogError(ex, $"error occured in {nameof(UpdateSMECardAsync)}");
            }
        }

        /// <summary>
        /// Remove mapping elements from ticket additional details and validate input values of type 'DateTime'.
        /// </summary>
        /// <param name="additionalDetails">Ticket addition details.</param>
        /// <param name="timeSpan">>Local time stamp.</param>
        /// <returns>Adaptive card item element json string.</returns>
        public static string ValidateAdditionalTicketDetails(string additionalDetails, TimeSpan timeSpan)
        {
            var details = JsonConvert.DeserializeObject<Dictionary<string, string>>(additionalDetails);

            RemoveMappingElement(details, "command");
            RemoveMappingElement(details, "teamId");
            RemoveMappingElement(details, "ticketId");
            RemoveMappingElement(details, "cardId");

            RemoveMappingElement(details, "Title");
            RemoveMappingElement(details, "Description");
            RemoveMappingElement(details, "RequestType");
            Dictionary<string, string> keyValuePair = new Dictionary<string, string>();
            if (details != null)
            {
                foreach (var item in details)
                {
                    try
                    {
                        keyValuePair.Add(item.Key, TicketHelper.ConvertToDateTimeoffset(DateTime.Parse(item.Value, CultureInfo.InvariantCulture), timeSpan).ToString(CultureInfo.InvariantCulture));
                    }
#pragma warning disable CA1031 // Do not catch general exception types
                    catch
#pragma warning restore CA1031 // Do not catch general exception types
                    {
                        keyValuePair.Add(item.Key, item.Value);
                    }
                }
            }

            return JsonConvert.SerializeObject(keyValuePair);
        }

        /// <summary>
        /// Converts json property to adaptive card element.
        /// </summary>
        /// <param name="elements">Adaptive item element Json object.</param>
        /// <returns>Returns adaptive card item element.</returns>
        public static List<AdaptiveElement> ConvertToAdaptiveCardItemElement(List<JObject> elements)
        {
            var adaptiveElements = new List<AdaptiveElement>();
            if (elements == null || elements.Count == 0)
            {
                return adaptiveElements;
            }

            foreach (var cardElement in elements)
            {
                var cardElementWithValues = JsonConvert.DeserializeObject<AdaptiveCardPlaceHolderMapper>(cardElement.ToString());

                switch (cardElementWithValues.InputType)
                {
                    case AdaptiveTextBlock.TypeName:
                        adaptiveElements.Add(AdaptiveElementHelper.ConvertToAdaptiveTextBlock(cardElement.ToString()));
                        break;
                    case AdaptiveTextInput.TypeName:
                        adaptiveElements.Add(AdaptiveElementHelper.ConvertToAdaptiveTextInput(cardElement.ToString()));
                        break;
                    case AdaptiveChoiceSetInput.TypeName:
                        adaptiveElements.Add(AdaptiveElementHelper.ConvertToAdaptiveChoiceSetInput(cardElement.ToString()));
                        break;
                    case AdaptiveDateInput.TypeName:
                        adaptiveElements.Add(AdaptiveElementHelper.ConvertToAdaptiveDateInput(cardElement.ToString()));
                        break;
                }
            }

            return adaptiveElements;
        }

        /// <summary>
        /// Convert json template to Adaptive card.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="cardTemplate">Adaptive card template.</param>
        /// <param name="showDateValidation">true if need to show validation message else false.</param>
        /// <param name="ticketDetails">Ticket details key value pair.</param>
        /// <returns>Adaptive card item element json string.</returns>
        public static List<AdaptiveElement> ConvertToAdaptiveCard(IStringLocalizer<Strings> localizer, string cardTemplate, bool showDateValidation, Dictionary<string, string> ticketDetails = null)
        {
            var cardTemplates = JsonConvert.DeserializeObject<List<JObject>>(cardTemplate);
            var cardTemplateElements = new List<JObject>();

            foreach (var template in cardTemplates)
            {
                var templateMapping = JsonConvert.DeserializeObject<AdaptiveCardPlaceHolderMapper>(template.ToString());
                if (templateMapping.InputType != "TextBlock")
                {
                    // get first observed display text if parsed from appSettings; rest all values will be set up directly in JSON payload.
                    if (templateMapping.Id == "IssueOccuredOn")
                    {
                        templateMapping.DisplayName = localizer.GetString("FirstObservedText");
                    }

                    // every input elements display name is integrated with the JSON payload
                    // and is converted to text block corresponding to input element
                    cardTemplateElements.Add(JObject.FromObject(new
                    {
                        type = AdaptiveTextBlock.TypeName,
                        displayName = templateMapping.DisplayName,
                        text = templateMapping.DisplayName,
                    }));

                    var templateMappingFieldValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(template.ToString());

                    if (ticketDetails != null)
                    {
                        templateMappingFieldValues["value"] = TryParseTicketDetailsKeyValuePair(ticketDetails, templateMapping.Id);
                    }

                    cardTemplateElements.Add(JObject.FromObject(templateMappingFieldValues));
                }
                else
                {
                    // Enabling validation message for First observed on date time field.
                    if (templateMapping.Id == "DateValidationMessage")
                    {
                        if (showDateValidation)
                        {
                            cardTemplateElements.Add(JObject.FromObject(new
                            {
                                type = AdaptiveTextBlock.TypeName,
                                id = "DateValidationMessage",
                                spacing = "None",
                                color = "Attention",
                                isVisible = "true",
                                text = localizer.GetString("DateValidationText"),
                            }));
                        }
                    }
                    else
                    {
                        cardTemplateElements.Add(template);
                    }
                }
            }

            // Parse and convert each elements to adaptive elements
            return ConvertToAdaptiveCardItemElement(cardTemplateElements);
        }

        /// <summary>
        /// Check and convert to DateTime adaptive text if input string is a valid date time.
        /// </summary>
        /// <param name="inputText">Input date time string.</param>
        /// <returns>Adaptive card supported date time format else return sting as-is.</returns>
        public static string AdaptiveTextParseWithDateTime(string inputText)
        {
            if (DateTime.TryParse(inputText, out DateTime inputDateTime))
            {
                return "{{DATE(" + inputDateTime.ToUniversalTime().ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture) + ", SHORT)}}";
            }

            return inputText;
        }

        /// <summary>
        /// Get values from dictionary.
        /// </summary>
        /// <param name="ticketDetails">Ticket additional details.</param>
        /// <param name="key">Dictionary key.</param>
        /// <returns>Dictionary value.</returns>
        public static string TryParseTicketDetailsKeyValuePair(Dictionary<string, string> ticketDetails, string key)
        {
            if (ticketDetails != null && ticketDetails.ContainsKey(key))
            {
                return ticketDetails[key];
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Remove item from dictionary.
        /// </summary>
        /// <param name="ticketDetails">Ticket details key value pair.</param>
        /// <param name="key">Dictionary key.</param>
        /// <returns>boolean value.</returns>
        public static bool RemoveMappingElement(Dictionary<string, string> ticketDetails, string key)
        {
            if (ticketDetails != null && ticketDetails.ContainsKey(key))
            {
                return ticketDetails.Remove(key);
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Get adaptive card column set.
        /// </summary>
        /// <param name="title">Column title.</param>
        /// <param name="value">Column value.</param>
        /// <returns>AdaptiveColumnSet.</returns>
        public static AdaptiveColumnSet GetAdaptiveCardColumnSet(string title, string value)
        {
            return new AdaptiveColumnSet
            {
                Columns = new List<AdaptiveColumn>
                {
                    new AdaptiveColumn
                    {
                        Width = "50",
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveTextBlock
                            {
                                HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                Text = $"{title}:",
                                Wrap = true,
                                Weight = AdaptiveTextWeight.Bolder,
                                Size = AdaptiveTextSize.Medium,
                            },
                        },
                    },
                    new AdaptiveColumn
                    {
                        Width = "100",
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveTextBlock
                            {
                                HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                Text = AdaptiveTextParseWithDateTime(value),
                                Wrap = true,
                            },
                        },
                    },
                },
            };
        }
    }
}