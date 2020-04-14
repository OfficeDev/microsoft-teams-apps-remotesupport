// <copyright file="TicketCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using AdaptiveCards;
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
        /// <param name="cardConfiguration">Card configuration.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="showValidationMessage">Represents whether to show validation message or not.</param>
        /// <param name="ticketDetail"> Information of the ticket which is being created.</param>
        /// <returns>Returns an attachment of new ticket.</returns>
        public static Attachment GetNewTicketCard(CardConfigurationEntity cardConfiguration, IStringLocalizer<Strings> localizer, bool showValidationMessage = false, TicketDetail ticketDetail = null)
        {
            cardConfiguration = cardConfiguration ?? throw new ArgumentNullException(nameof(cardConfiguration));

            string issueTitle = string.Empty;
            string issueDescription = string.Empty;

            var dynamicElements = new List<AdaptiveElement>();
            var ticketAdditionalFields = new List<AdaptiveElement>();
            bool showTitleValidation = false;
            bool showDescriptionValidation = false;
            bool showDateValidation = false;

            if (showValidationMessage)
            {
                ticketDetail = ticketDetail ?? throw new ArgumentNullException(nameof(ticketDetail));
                if (string.IsNullOrWhiteSpace(ticketDetail.Title))
                {
                    showTitleValidation = true;
                }
                else
                {
                    issueTitle = ticketDetail.Title;
                }

                if (string.IsNullOrWhiteSpace(ticketDetail.Description))
                {
                    showDescriptionValidation = true;
                }
                else
                {
                    issueDescription = ticketDetail.Description;
                }

                if (ticketDetail.IssueOccuredOn == null || DateTimeOffset.Compare(ticketDetail.IssueOccuredOn, DateTime.Today) > 0 || string.IsNullOrEmpty(ticketDetail.IssueOccuredOn.ToString(CultureInfo.InvariantCulture)))
                {
                    showDateValidation = true;
                }
            }

            ticketAdditionalFields = CardHelper.ConvertToAdaptiveCard(localizer, cardConfiguration.CardTemplate, showDateValidation);

            dynamicElements.AddRange(new List<AdaptiveElement>
            {
                new AdaptiveTextBlock
                {
                    Text = localizer.GetString("NewRequestTitle"),
                    Weight = AdaptiveTextWeight.Bolder,
                    Size = AdaptiveTextSize.Large,
                },
                new AdaptiveTextBlock()
                {
                    Text = localizer.GetString("TellUsAboutProblemText"),
                    Spacing = AdaptiveSpacing.Small,
                },
                new AdaptiveTextBlock()
                {
                    Text = localizer.GetString("TitleDisplayText"),
                    Spacing = AdaptiveSpacing.Medium,
                },
                new AdaptiveTextInput()
                {
                    Id = "Title",
                    MaxLength = 100,
                    Placeholder = localizer.GetString("TitlePlaceHolderText"),
                    Spacing = AdaptiveSpacing.Small,
                    Value = issueTitle,
                },
                new AdaptiveTextBlock()
                {
                    Text = localizer.GetString("TitleValidationText"),
                    Spacing = AdaptiveSpacing.None,
                    IsVisible = showTitleValidation,
                    Color = AdaptiveTextColor.Attention,
                },
                new AdaptiveTextBlock()
                {
                    Text = localizer.GetString("DescriptionText"),
                    Spacing = AdaptiveSpacing.Medium,
                },
                new AdaptiveTextInput()
                {
                    Id = "Description",
                    MaxLength = 500,
                    IsMultiline = true,
                    Placeholder = localizer.GetString("DesciptionPlaceHolderText"),
                    Spacing = AdaptiveSpacing.Small,
                    Value = issueDescription,
                },
                new AdaptiveTextBlock()
                {
                    Text = localizer.GetString("DescriptionValidationText"),
                    Spacing = AdaptiveSpacing.None,
                    IsVisible = showDescriptionValidation,
                    Color = AdaptiveTextColor.Attention,
                },
                new AdaptiveTextBlock()
                {
                    Text = localizer.GetString("RequestTypeText"),
                    Spacing = AdaptiveSpacing.Medium,
                },
                new AdaptiveChoiceSetInput
                {
                    Choices = new List<AdaptiveChoice>
                    {
                        new AdaptiveChoice
                        {
                            Title = localizer.GetString("NormalText"),
                            Value = localizer.GetString("NormalText"),
                        },
                        new AdaptiveChoice
                        {
                            Title = localizer.GetString("UrgentText"),
                            Value = localizer.GetString("UrgentText"),
                        },
                    },
                    Id = "RequestType",
                    Value = !string.IsNullOrEmpty(ticketDetail?.RequestType) ? ticketDetail?.RequestType : localizer.GetString("NormalText"),
                    Style = AdaptiveChoiceInputStyle.Expanded,
                },
            });

            dynamicElements.AddRange(ticketAdditionalFields);

            AdaptiveCard ticketDetailsPersonalChatCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = dynamicElements,
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("SendRequestButtonText"),
                        Id = "SendRequest",
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = Constants.MessageBackActionType,
                                Text = Constants.SendRequestAction,
                            },
                            CardId = cardConfiguration?.CardId,
                            TeamId = cardConfiguration?.TeamId,
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
            foreach (KeyValuePair<string, string> item in ticketAdditionalDetail)
            {
                ticketAdditionalFields.Add(CardHelper.GetAdaptiveCardColumnSet(item.Key, item.Value));
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
