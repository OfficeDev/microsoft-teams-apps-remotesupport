// <copyright file="SearchHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RemoteSupport.Common;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Providers;

    /// <summary>
    /// Class that handles the search activities for messaging extension.
    /// </summary>
    public static class SearchHelper
    {
        /// <summary>
        /// Truncate the length of description to show in thumbnail card.
        /// </summary>
        private const int TruncateDescriptionLength = 50;

        /// <summary>
        /// Use Ellipsis for a long description.
        /// </summary>
        private const string Ellipsis = "...";

        /// <summary>
        /// Search text parameter name defined in the application manifest file.
        /// </summary>
        private const string SearchTextParameterName = "searchText";

        /// <summary>
        /// Feedback - text that renders share feedback card.
        /// </summary>
        private const string GoToOriginalThreadUrl = "https://teams.microsoft.com/l/message/";

        /// <summary>
        /// Get the value of the searchText parameter in the messaging extension query.
        /// </summary>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <returns>A value of the searchText parameter.</returns>
        public static string GetSearchQueryString(MessagingExtensionQuery query)
        {
            var messageExtensionInputText = query?.Parameters.FirstOrDefault(parameter => parameter.Name.Equals(SearchTextParameterName, StringComparison.OrdinalIgnoreCase));
            return messageExtensionInputText?.Value?.ToString();
        }

        /// <summary>
        /// Get the results from Azure search service and populate the result (card + preview).
        /// </summary>
        /// <param name="query">Query which the user had typed in message extension search.</param>
        /// <param name="commandId">Command id to determine which tab in message extension has been invoked.</param>
        /// <param name="count">Count for pagination.</param>
        /// <param name="skip">Skip for pagination.</param>
        /// <param name="searchService">Search service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="requestorId">Requester id of the user to get specific tickets.</param>
        /// <param name="onCallSMEUsers">OncallSMEUsers to give support from group-chat or on-call.</param>
        /// <returns><see cref="Task"/> Returns MessagingExtensionResult which will be used for providing the card.</returns>
        public static async Task<MessagingExtensionResult> GetSearchResultAsync(
            string query,
            string commandId,
            int? count,
            int? skip,
            ITicketSearchService searchService,
            IStringLocalizer<Strings> localizer,
            string requestorId = "",
            string onCallSMEUsers = "")
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            IList<TicketDetail> searchServiceResults;

            // commandId should be equal to Id mentioned in Manifest file under composeExtensions section.
            switch (commandId)
            {
                case Constants.UrgentCommandId:
                    searchServiceResults = await searchService?.SearchTicketsAsync(TicketSearchScope.UrgentTickets, query, count, skip);
                    composeExtensionResult = GetMessagingExtensionResult(searchServiceResults, localizer, commandId);
                    break;

                case Constants.AssignedCommandId:
                    searchServiceResults = await searchService?.SearchTicketsAsync(TicketSearchScope.AssignedTickets, query, count, skip);
                    composeExtensionResult = GetMessagingExtensionResult(searchServiceResults, localizer, commandId);
                    break;

                case Constants.UnassignedCommandId:
                    searchServiceResults = await searchService?.SearchTicketsAsync(TicketSearchScope.UnassignedTickets, query, count, skip);
                    composeExtensionResult = GetMessagingExtensionResult(searchServiceResults, localizer, commandId);
                    break;

                case Constants.ActiveCommandId:
                    searchServiceResults = await searchService?.SearchTicketsAsync(TicketSearchScope.ActiveTickets, query, count, skip, requestorId);
                    composeExtensionResult = GetMessagingExtensionResult(searchServiceResults, localizer, commandId, onCallSMEUsers);
                    break;

                case Constants.ClosedCommandId:
                    searchServiceResults = await searchService?.SearchTicketsAsync(TicketSearchScope.ClosedTickets, query, count, skip, requestorId);
                    composeExtensionResult = GetMessagingExtensionResult(searchServiceResults, localizer, commandId);
                    break;
            }

            return composeExtensionResult;
        }

        /// <summary>
        /// Get result for messaging extension tab.
        /// </summary>
        /// <param name="searchServiceResults">List of tickets from Azure search service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="commandId">Command id to determine which tab in message extension has been invoked.</param>
        /// <param name="onCallSMEUsers">OncallSMEUsers to give support from group-chat or on-call.</param>
        /// <returns><see cref="Task"/> Returns MessagingExtensionResult which will be shown in messaging extension tab.</returns>
        public static MessagingExtensionResult GetMessagingExtensionResult(
            IList<TicketDetail> searchServiceResults,
            IStringLocalizer<Strings> localizer,
            string commandId = "",
            string onCallSMEUsers = "")
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            if (searchServiceResults != null)
            {
                foreach (var ticket in searchServiceResults)
                {
                    var dynamicElements = new List<AdaptiveElement>
                    {
                        CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("RequestNumberText"), $"#{ticket.TicketId}"),
                        CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("TitleDisplayText"), ticket.Title),
                        CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("DescriptionText"), ticket.Description),
                        CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("CreatedOnText"), ticket.CreatedOn.ToString(CultureInfo.InvariantCulture)),
                    };

                    AdaptiveCard commandIdCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
                    {
                        Body = dynamicElements,
                        Actions = new List<AdaptiveAction>(),
                    };

                    if (commandId == Constants.ActiveCommandId && !string.IsNullOrEmpty(onCallSMEUsers))
                    {
                        commandIdCard.Actions.Add(
                            new AdaptiveOpenUrlAction
                            {
                                Title = localizer.GetString("EscalateButtonText"),
                                Url = new Uri(CreateGroupChat(onCallSMEUsers, ticket.TicketId, ticket.RequesterName, localizer)),
                            });
                    }
                    else if ((commandId == Constants.UrgentCommandId || commandId == Constants.AssignedCommandId || commandId == Constants.UnassignedCommandId) && ticket.SmeConversationId != null)
                    {
                        commandIdCard.Actions.Add(
                            new AdaptiveOpenUrlAction
                            {
                                Title = localizer.GetString("GoToOriginalThreadButtonText"),
                                Url = new Uri(CreateDeepLinkToThread(ticket.SmeConversationId)),
                            });
                    }

                    ThumbnailCard previewCard = new ThumbnailCard
                    {
                        Title = $"<b>{HttpUtility.HtmlEncode(ticket.Title)} | {HttpUtility.HtmlEncode(ticket.Severity == (int)TicketSeverity.Urgent ? localizer.GetString("UrgentText") : localizer.GetString("NormalText"))}</b>",
                        Subtitle = ticket.Description.Length <= TruncateDescriptionLength ? HttpUtility.HtmlEncode(ticket.Description) : HttpUtility.HtmlEncode(ticket.Description.Substring(0, 45)) + Ellipsis,
                        Text = ticket.RequesterName,
                    };
                    composeExtensionResult.Attachments.Add(new Attachment
                    {
                        ContentType = AdaptiveCard.ContentType,
                        Content = commandIdCard,
                    }.ToMessagingExtensionAttachment(previewCard.ToAttachment()));
                }
            }

            return composeExtensionResult;
        }

        /// <summary>
        /// Returns go to original thread Uri which will help in opening the original conversation about the ticket.
        /// </summary>
        /// <param name="threadConversationId">The thread along with message Id stored in storage table.</param>
        /// <returns>Original thread Uri.</returns>
        private static string CreateDeepLinkToThread(string threadConversationId)
        {
            string[] threadAndMessageId = threadConversationId.Split(";");
            var threadId = threadAndMessageId[0];
            var messageId = threadAndMessageId[1].Split("=")[1];
            return $"{GoToOriginalThreadUrl}{threadId}/{messageId}";
        }

        /// <summary>
        /// Returns the group chat Uri which will help to create group chat with on calls SME users.
        /// </summary>
        /// <param name="onCallSMENames">The on-call SME users which are supported for group chat.</param>
        /// <param name="ticketId">Ticket id of the request.</param>
        /// <param name="requesterName">Requester name of the ticket.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Group chat Uri.</returns>
        private static string CreateGroupChat(string onCallSMENames, string ticketId, string requesterName, IStringLocalizer<Strings> localizer)
        {
            var groupChatTitle = localizer.GetString("GroupName", ticketId);
            var previewText = localizer.GetString("MessageContent", requesterName);
            return $"https://teams.microsoft.com/l/chat/0/0?users={Uri.EscapeDataString(onCallSMENames)}&topicName={Uri.EscapeDataString(groupChatTitle)}&message={Uri.EscapeDataString(previewText)}";
        }
    }
}
