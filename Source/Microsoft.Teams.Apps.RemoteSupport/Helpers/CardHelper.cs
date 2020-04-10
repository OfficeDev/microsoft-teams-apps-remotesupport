// <copyright file="CardHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Hosting;
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
        /// <param name="taskModuleRequestData">Task module invoke request value payload.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="activityId">Task module activity Id.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns task module response.</returns>
        public static TaskModuleResponse GetTaskModuleResponse(string applicationBasePath, string customAPIAuthenticationToken, TaskModuleRequest taskModuleRequestData, TelemetryClient telemetryClient, string activityId, IStringLocalizer<Strings> localizer)
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
        /// <param name="environment">Current environment.</param>
        /// <param name="cardConfigurationStorageProvider">Card configuration.</param>
        /// <param name="ticketDetail">Details of the ticket to be edited.</param>
        /// <param name="showValidationMessage">Determines whether to show validation message on screen or not.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="existingTicketDetail">Existing ticket details.</param>
        /// <returns>Returns edit ticket adaptive card.</returns>
        public static TaskModuleResponse GetEditTicketAdaptiveCard(IHostingEnvironment environment, ICardConfigurationStorageProvider cardConfigurationStorageProvider, TicketDetail ticketDetail, bool showValidationMessage, IStringLocalizer<Strings> localizer, TicketDetail existingTicketDetail = null)
        {
            var cardTemplate = cardConfigurationStorageProvider?.GetConfigurationsByCardIdAsync(ticketDetail?.CardId).Result;
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = EditRequestCard.GetEditRequestCard(environment, ticketDetail, cardTemplate, localizer, showValidationMessage, existingTicketDetail),
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
        /// <param name="logger">application logger.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>task that updates card.</returns>
        public static async Task UpdateSMECardAsync(ITurnContext turnContext, TicketDetail ticketDetail, IMessageActivity messageActivity, string applicationBasePath, IStringLocalizer<Strings> localizer, ILogger<RemoteSupportActivityHandler> logger, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            messageActivity = messageActivity ?? throw new ArgumentNullException(nameof(messageActivity));
            try
            {
                // Update the card in the SME team.
                var updateCardActivity = new Activity(ActivityTypes.Message)
                {
                    Id = ticketDetail?.SmeTicketActivityId,
                    Conversation = new ConversationAccount { Id = ticketDetail.SmeConversationId },
                    Attachments = new List<Attachment> { new SmeTicketCard(ticketDetail).GetTicketDetailsForSMEChatCard(ticketDetail, applicationBasePath, localizer) },
                };

                messageActivity.Conversation = new ConversationAccount { Id = ticketDetail.SmeConversationId };

                await turnContext.Adapter.UpdateActivityAsync(turnContext, updateCardActivity, cancellationToken);
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
        /// Convert json string to adaptive card attachment.
        /// </summary>
        /// <param name="cardPayload">Card json content.</param>
        /// <returns>An attachment with required card content.</returns>
        public static Attachment ConvertPayloadToAttachment(string cardPayload)
        {
            AdaptiveCard card = AdaptiveCard.FromJson(cardPayload).Card;
            return new Attachment()
            {
                Content = card,
                ContentType = AdaptiveCard.ContentType,
            };
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
            RemoveMappingElement(details, "TeamId");
            RemoveMappingElement(details, "ticketId");
            RemoveMappingElement(details, "CardId");
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
        /// Convert json template to Adaptive card item element.
        /// </summary>
        /// <param name="cardTemplate">Adaptive card template.</param>
        /// <returns>Adaptive card item element json string.</returns>
        public static string ConvertToAdaptiveCardItemElement(string cardTemplate)
        {
            var jsonObjects = JsonConvert.DeserializeObject<List<object>>(cardTemplate);
            var tempElements = new List<object>();
            foreach (var item in jsonObjects)
            {
                var mapping = JsonConvert.DeserializeObject<AdaptiveCardPlaceHolderMapper>(item.ToString());
                if (mapping.InputType != "TextBlock")
                {
                    var displayName = "{\"type\":\"TextBlock\",\"text\":\"" + mapping.Id.Split('_')[0] + "\"}";
                    tempElements.Add(JsonConvert.DeserializeObject(displayName));
                }

                tempElements.Add(item);
            }

            return JsonConvert.SerializeObject(tempElements).TrimStart('[').TrimEnd(']');
        }

        /// <summary>
        /// Convert json template to Adaptive card edit item element.
        /// </summary>
        /// <param name="cardTemplate">Adaptive card template.</param>
        /// <param name="ticketDetails">Ticket details Key value pair.</param>
        /// <returns>Adaptive card item element json string.</returns>
        public static string ConvertToAdaptiveCardEditItemElement(string cardTemplate, Dictionary<string, string> ticketDetails)
        {
            var jsonObjects = JsonConvert.DeserializeObject<List<object>>(cardTemplate);
            var tempElements = new List<object>();
            foreach (var item in jsonObjects)
            {
                var mapping = JsonConvert.DeserializeObject<AdaptiveCardPlaceHolderMapper>(item.ToString());
                if (mapping.InputType != "TextBlock")
                {
                    var displayName = "{\"type\":\"TextBlock\",\"text\":\"" + mapping.Id.Split('_')[0] + "\"}";
                    var mappingValueField = JsonConvert.DeserializeObject<Dictionary<string, object>>(item.ToString());
                    if (!mappingValueField.ContainsKey("value"))
                    {
                        mappingValueField.Add("value", GetDictionaryValue(ticketDetails, mapping.Id));
                    }
                    else
                    {
                        mappingValueField["value"] = GetDictionaryValue(ticketDetails, mapping.Id);
                    }

                    tempElements.Add(JsonConvert.DeserializeObject(displayName));
                    tempElements.Add(JsonConvert.DeserializeObject(JsonConvert.SerializeObject(mappingValueField)));
                }
                else
                {
                    tempElements.Add(item);
                }
            }

            return JsonConvert.SerializeObject(tempElements).TrimStart('[').TrimEnd(']');
        }

        /// <summary>
        /// Convert Date time format to adaptive card text feature.
        /// </summary>
        /// <param name="inputText">Input date time string.</param>
        /// <returns>Adaptive card supported date time format.</returns>
        public static string FormatDateStringToAdaptiveCardDateFormat(string inputText)
        {
            try
            {
                return "{{DATE(" + DateTime.Parse(inputText, CultureInfo.InvariantCulture).ToUniversalTime().ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture) + ", SHORT)}}";
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch
#pragma warning restore CA1031 // Do not catch general exception types
            {
                return inputText;
            }
        }

        /// <summary>
        /// Get values from dictionary.
        /// </summary>
        /// <param name="ticketDetails">Ticket additional details.</param>
        /// <param name="key">Dictionary key.</param>
        /// <returns>Dictionary value.</returns>
        public static string GetDictionaryValue(Dictionary<string, string> ticketDetails, string key)
        {
            if (ticketDetails != null && ticketDetails.ContainsKey(key))
            {
                return EscapeCharactersInString(ticketDetails[key]);
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
        /// Replace template parameters.
        /// </summary>
        /// <param name="input">input key.</param>
        /// <param name="variablesToValues">key value pair.</param>
        /// <returns>return valid string.</returns>
        public static string ResolveTemplateParams(string input, Dictionary<string, string> variablesToValues)
        {
            string output = input?.Substring(0);

            if (variablesToValues != null)
            {
                foreach (KeyValuePair<string, string> kvp in variablesToValues)
                {
                    output = output.Replace($"_{kvp.Key}_", kvp.Value, StringComparison.OrdinalIgnoreCase);
                }
            }

            return output;
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
                                Text = FormatDateStringToAdaptiveCardDateFormat(value),
                                Wrap = true,
                            },
                        },
                    },
                },
            };
        }

        /// <summary>
        /// Replace double quotes and escape characters in string.
        /// </summary>
        /// <param name="inputString">Input string which needs to be validated.</param>
        /// <returns>Returns valid string after escaping characters.</returns>
        private static string EscapeCharactersInString(string inputString)
        {
            if (string.IsNullOrWhiteSpace(inputString))
            {
                return string.Empty;
            }

            inputString = JsonConvert.SerializeObject(inputString);
            var match = Regex.Match(inputString, "^[\"](.*)\"");
            return match.Success ? match.Groups[1].Value : string.Empty;
        }
    }
}