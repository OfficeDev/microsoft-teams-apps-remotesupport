// <copyright file="TicketCard.cs" company="Microsoft">
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
    using Microsoft.Teams.Apps.RemoteSupport.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// Provides adaptive cards for creating and editing new ticket information.
    /// </summary>
    public static class TicketCard
    {
        /// <summary>
        /// Get the create new ticket card.
        /// </summary>
        /// <param name="environment">Current environment.</param>
        /// <param name="cardConfiuration">Card configuration.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="showValidationMessage">Represents whether to show validation message or not.</param>
        /// <param name="ticketDetail"> Information of the ticket which is being created.</param>
        /// <param name="ticketAdditionalDetails">Additional information of the ticket which is being created.</param>
        /// <returns>Returns an attachment of new ticket.</returns>
        public static Attachment GetNewTicketCard(IHostingEnvironment environment, CardConfigurationEntity cardConfiuration, IStringLocalizer<Strings> localizer, bool showValidationMessage = false, TicketDetail ticketDetail = null, string ticketAdditionalDetails = null)
        {
            environment = environment ?? throw new ArgumentNullException(nameof(environment));
            string showTitleValidation = "false";
            string showDescriptionValidation = "false";
            string showDateValidation = "false";
            string issueTitle = string.Empty;
            string issueDescription = string.Empty;
            string issueDateString = DateTime.Now.ToString(CultureInfo.InvariantCulture);
            string dynamicTemplate;

            if (showValidationMessage)
            {
                ticketDetail = ticketDetail ?? throw new ArgumentNullException(nameof(ticketDetail));
                if (string.IsNullOrWhiteSpace(ticketDetail.Title))
                {
                    showTitleValidation = "true";
                }
                else
                {
                    issueTitle = ticketDetail.Title;
                }

                if (string.IsNullOrWhiteSpace(ticketDetail.Description))
                {
                    showDescriptionValidation = "true";
                }
                else
                {
                    issueDescription = ticketDetail.Description;
                }

                if (ticketDetail.IssueOccuredOn == null || DateTimeOffset.Compare(ticketDetail.IssueOccuredOn, DateTime.Today) > 0 || string.IsNullOrEmpty(ticketDetail.IssueOccuredOn.ToString(CultureInfo.InvariantCulture)))
                {
                    showDateValidation = "true";
                }
                else
                {
                    issueDateString = ticketDetail.IssueOccuredOn.ToString(CultureInfo.InvariantCulture);
                }

                var ticketAdditionalDetail = JsonConvert.DeserializeObject<Dictionary<string, string>>(ticketAdditionalDetails);
                dynamicTemplate = CardHelper.ConvertToAdaptiveCardEditItemElement(cardConfiuration?.CardTemplate, ticketAdditionalDetail);
            }
            else
            {
                dynamicTemplate = CardHelper.ConvertToAdaptiveCardItemElement(cardConfiuration?.CardTemplate);
            }

            string cardJsonFilePath = Path.Combine(environment.ContentRootPath, ".\\Cards\\NewTicket.json");
            string cardPayload = File.ReadAllText(cardJsonFilePath);

            Dictionary<string, string> variablesToValues = new Dictionary<string, string>()
            {
                { "DynamicContent", dynamicTemplate },
                { "issueTitle", issueTitle },
                { "issueDescription", issueDescription },
                { "severity", localizer.GetString("NormalText") },
                { "issueDate", issueDateString },
                { "titleValidationText", localizer.GetString("TitleValidationText") },
                { "showTitleValidation", showTitleValidation },
                { "descriptionValidationText", localizer.GetString("DescriptionValidationText") },
                { "showDescriptionValidation", showDescriptionValidation },
                { "dateValidationText", localizer.GetString("DateValidationText") },
                { "showDateValidation", showDateValidation },
                { "maxDate", DateTime.Now.ToString("YYYY-MM-DD", CultureInfo.InvariantCulture) },
                { "cardId", cardConfiuration.CardId },
                { "teamId", cardConfiuration.TeamId },
                { "NewRequestTitle", localizer.GetString("NewRequestTitle") },
                { "TellUsAboutProblemText", localizer.GetString("TellUsAboutProblemText") },
                { "TitleDisplayText", localizer.GetString("TitleDisplayText") },
                { "TitlePlaceHolderText", localizer.GetString("TitlePlaceHolderText") },
                { "DescriptionText", localizer.GetString("DescriptionText") },
                { "DesciptionPlaceHolderText", localizer.GetString("DesciptionPlaceHolderText") },
                { "RequestTypeText", localizer.GetString("RequestTypeText") },
                { "NormalText", localizer.GetString("NormalText") },
                { "UrgentText", localizer.GetString("UrgentText") },
                { "SendRequestButtonText", localizer.GetString("SendRequestButtonText") },
            };

            // Removing extra character from json stream.
            cardPayload = CardHelper.ResolveTemplateParams(cardPayload, variablesToValues).TrimStart('[').TrimEnd(']');
            return CardHelper.ConvertPayloadToAttachment(cardPayload);
        }

        /// <summary>
        /// Card to show ticket details in 1:1 chat with bot after submitting request details.
        /// </summary>
        /// <param name="ticketDetail">New ticket values entered by user.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="isEdited">flag that sets when card is edited.</param>
        /// <returns>An attachment with ticket details.</returns>
        public static Attachment GetTicketDetailsForPersonalChatCard(TicketDetail ticketDetail, IStringLocalizer<Strings> localizer, bool isEdited = false)
        {
            ticketDetail = ticketDetail ?? throw new ArgumentNullException(nameof(ticketDetail));
            Dictionary<string, string> ticketAdditionalDetail = JsonConvert.DeserializeObject<Dictionary<string, string>>(ticketDetail.AdditionalProperties);
            var dynamicElements = new List<AdaptiveElement>();
            var ticketAdditionalFields = new List<AdaptiveElement>();

            CardHelper.RemoveMappingElement(ticketAdditionalDetail, "CardId");
            CardHelper.RemoveMappingElement(ticketAdditionalDetail, "TeamId");
            CardHelper.RemoveMappingElement(ticketAdditionalDetail, "Title");
            CardHelper.RemoveMappingElement(ticketAdditionalDetail, "Description");
            CardHelper.RemoveMappingElement(ticketAdditionalDetail, "RequestType");

            foreach (KeyValuePair<string, string> item in ticketAdditionalDetail)
            {
                ticketAdditionalFields.Add(CardHelper.GetAdaptiveCardColumnSet(item.Key.Split("_")?[0], item.Value));
            }

            dynamicElements.AddRange(new List<AdaptiveElement>
            {
                new AdaptiveTextBlock
                {
                    Text = isEdited == true ? localizer.GetString("RequestUpdatedText") : localizer.GetString("RequestSubmittedText"),
                    Weight = AdaptiveTextWeight.Bolder,
                    Size = AdaptiveTextSize.Large,
                },
                new AdaptiveTextBlock()
                {
                    Text = localizer.GetString("RequestSubmittedContent"),
                    Wrap = true,
                    Spacing = AdaptiveSpacing.None,
                },
                CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("RequestNumberText"), $"#{ticketDetail.RowKey}"),
                CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("RequestTypeText"), ticketDetail.RequestType),
                CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("TitleDisplayText"), ticketDetail.Title),
            });
            dynamicElements.AddRange(ticketAdditionalFields);
            dynamicElements.Add(CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("DescriptionText"), ticketDetail.Description));

            AdaptiveCard ticketDetailsPersonalChatCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = dynamicElements,
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("EditTicketActionText"),
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = Constants.FetchActionType,
                            },
                            Command = Constants.EditRequestAction,
                            PostedValues = ticketDetail.TicketId,
                        },
                    },
                    new AdaptiveShowCardAction()
                    {
                        Title = localizer.GetString("WithdrawRequestActionText"),
                        Card = WithdrawCard.ConfirmationCard(ticketDetail.TicketId, localizer),
                    },
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("NewRequestButtonText"),
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = Constants.MessageBackActionType,
                                Text = Constants.NewRequestAction,
                            },
                        },
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = ticketDetailsPersonalChatCard,
            };
        }
    }
}
