// <copyright file="EditRequestCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using AdaptiveCards;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RemoteSupport.Common;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.Teams.Apps.RemoteSupport.Helpers;
    using Newtonsoft.Json;

    /// <summary>
    /// Class holds card for Edit request.
    /// </summary>
    public static class EditRequestCard
    {
        /// <summary>
        /// Gets Edit card for task module.
        /// </summary>
        /// <param name="environment">Current environment.</param>
        /// <param name="ticketDetail">Ticket details from user.</param>
        /// <param name="cardConfiuration">Card configuration.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="showValidationMessage">show ValidationMessage.</param>
        /// <param name="existingTicketDetail">Existing ticket details.</param>
        /// <returns>Returns an attachment of edit card.</returns>
        public static Attachment GetEditRequestCard(IHostingEnvironment environment, TicketDetail ticketDetail, CardConfigurationEntity cardConfiuration, IStringLocalizer<Strings> localizer, bool showValidationMessage = false, TicketDetail existingTicketDetail = null)
        {
            ticketDetail = ticketDetail ?? throw new ArgumentNullException(nameof(ticketDetail));
            string showTitleValidation = "false";
            string showDescriptionValidation = "false";
            string showDateValidation = "false";
            string issueDateString = DateTime.Now.ToString(CultureInfo.InvariantCulture);

            var editCardJsonFilePath = Path.Combine(environment?.ContentRootPath, ".\\Cards\\EditTicket.json");
            var cardPayload = File.ReadAllText(editCardJsonFilePath);
            var ticketAdditionalDetails = JsonConvert.DeserializeObject<Dictionary<string, string>>(ticketDetail.AdditionalProperties);
            string dynamicTemplate = CardHelper.ConvertToAdaptiveCardEditItemElement(cardConfiuration?.CardTemplate, ticketAdditionalDetails);

            if (showValidationMessage)
            {
                if (string.IsNullOrWhiteSpace(ticketDetail.Title))
                {
                    showTitleValidation = "true";
                }

                if (string.IsNullOrWhiteSpace(ticketDetail.Description))
                {
                    showDescriptionValidation = "true";
                }

                if (DateTimeOffset.Compare(ticketDetail.IssueOccuredOn, DateTime.Today) > 0 || string.IsNullOrEmpty(ticketDetail.IssueOccuredOn.ToString(CultureInfo.InvariantCulture)))
                {
                    showDateValidation = "true";
                }
                else if (existingTicketDetail != null && DateTimeOffset.Compare(existingTicketDetail.IssueOccuredOn, ticketDetail.IssueOccuredOn) < 0)
                {
                    showDateValidation = "true";
                }
                else
                {
                    issueDateString = ticketDetail.IssueOccuredOn.ToString(CultureInfo.InvariantCulture);
                }
            }

            Dictionary<string, string> variablesToValues = new Dictionary<string, string>()
            {
                { "DynamicContent", dynamicTemplate },
                { "issueTitle", CardHelper.GetDictionaryValue(ticketAdditionalDetails, "Title") },
                { "issueDescription", CardHelper.GetDictionaryValue(ticketAdditionalDetails, "Description") },
                { "issueDate", issueDateString },
                { "titleValidationText", localizer.GetString("TitleValidationText") },
                { "showTitleValidation", showTitleValidation },
                { "descriptionValidationText", localizer.GetString("DescriptionValidationText") },
                { "showDescriptionValidation", showDescriptionValidation },
                { "dateValidationText", localizer.GetString("DateValidationText") },
                { "showDateValidation", showDateValidation },
                { "maxDate", DateTime.Now.ToString("YYYY-DD-MM", CultureInfo.InvariantCulture) },
                { "cardId", cardConfiuration.CardId },
                { "teamId", cardConfiuration.TeamId },
                { "issueRequestType", ticketDetail.RequestType },
                { "ticketId", ticketDetail.TicketId },
                { "TitleDisplayText", localizer.GetString("TitleDisplayText") },
                { "TitlePlaceHolderText", localizer.GetString("TitlePlaceHolderText") },
                { "DescriptionText", localizer.GetString("DescriptionText") },
                { "DesciptionPlaceHolderText", localizer.GetString("DesciptionPlaceHolderText") },
                { "RequestTypeText", localizer.GetString("RequestTypeText") },
                { "NormalText", localizer.GetString("NormalText") },
                { "UrgentText", localizer.GetString("UrgentText") },
                { "UpdateActionText", localizer.GetString("UpdateActionText") },
                { "CancelButtonText", localizer.GetString("CancelButtonText") },
            };

            cardPayload = CardHelper.ResolveTemplateParams(cardPayload, variablesToValues);
            AdaptiveCard card = AdaptiveCard.FromJson(cardPayload).Card;
            return new Attachment()
            {
                Content = card,
                ContentType = AdaptiveCard.ContentType,
            };
        }

        /// <summary>
        /// Construct the card to render error message text to task module.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Card attachment.</returns>
        public static Attachment GetClosedErrorCard(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard closedErrorCard = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("ClosedErrorMessage"),
                        Wrap = true,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = closedErrorCard,
            };
        }
    }
}
