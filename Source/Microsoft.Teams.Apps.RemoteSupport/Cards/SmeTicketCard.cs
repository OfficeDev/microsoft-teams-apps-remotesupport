// <copyright file="SmeTicketCard.cs" company="Microsoft">
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
    using Newtonsoft.Json;

    /// <summary>
    /// Represents an SME ticket used for both in place card update activity within SME channel
    /// when changing the ticket status and notification card when bot posts user question to SME channel.
    /// </summary>
    public class SmeTicketCard
    {
        private readonly TicketDetail ticket;

        /// <summary>
        /// Initializes a new instance of the <see cref="SmeTicketCard"/> class.
        /// </summary>
        /// <param name="ticket">The ticket model with the latest details.</param>
        public SmeTicketCard(TicketDetail ticket)
        {
            this.ticket = ticket;
        }

        /// <summary>
        /// Returns an attachment based on the state and information of the ticket.
        /// </summary>
        /// <param name="cardElementMapping">Represents Adaptive card item element {Id, display name} mapping.</param>
        /// <param name="ticketDetail"> ticket values entered by user.</param>
        /// <param name="applicationBasePath">Represents the Application base URI.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns the attachment that will be sent in a message.</returns>
        public Attachment GetTicketDetailsForSMEChatCard(Dictionary<string, string> cardElementMapping, TicketDetail ticketDetail, string applicationBasePath, IStringLocalizer<Strings> localizer)
        {
            ticketDetail = ticketDetail ?? throw new ArgumentNullException(nameof(ticketDetail));
            cardElementMapping = cardElementMapping ?? throw new ArgumentNullException(nameof(cardElementMapping));

            Dictionary<string, string> ticketAdditionalDetail = JsonConvert.DeserializeObject<Dictionary<string, string>>(ticketDetail.AdditionalProperties);
            var dynamicElements = new List<AdaptiveElement>();
            var ticketAdditionalFields = new List<AdaptiveElement>();

            foreach (KeyValuePair<string, string> ticketField in ticketAdditionalDetail)
            {
                string key = ticketField.Key;

                // Issue occured on text block name needs to be fetched from card templates
                // here IssueOccurredOn is the id of text block
                if (ticketField.Key.Equals(CardConstants.IssueOccurredOnId, StringComparison.OrdinalIgnoreCase))
                {
                    key = localizer.GetString("FirstObservedText");
                }

                ticketAdditionalFields.Add(CardHelper.GetAdaptiveCardColumnSet(cardElementMapping.ContainsKey(key) ? cardElementMapping[key] : key, ticketField.Value));
            }

            dynamicElements.AddRange(new List<AdaptiveElement>
            {
                new AdaptiveColumnSet
                {
                    Columns = new List<AdaptiveColumn>
                    {
                        new AdaptiveColumn
                        {
                            Width = "12",
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                    Text = localizer.GetString("RequestDetailsText"),
                                    Weight = AdaptiveTextWeight.Bolder,
                                    Size = AdaptiveTextSize.Large,
                                },
                            },
                        },
                        new AdaptiveColumn
                        {
                            Width = "3",
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveImage
                                {
                                    Url = new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Artifacts/Urgent.png", applicationBasePath?.Trim('/'))),
                                    Size = AdaptiveImageSize.Large,
                                    AltText = localizer.GetString("UrgentText"),
                                    IsVisible = ticketDetail.RequestType == Constants.UrgentString,
                                },
                            },
                        },
                    },
                },
                new AdaptiveTextBlock()
                {
                    Text = localizer.GetString("SmeRequestDetailText", this.ticket.RequesterName),
                    Spacing = AdaptiveSpacing.None,
                },
                CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("RequestNumberText"), $"#{ticketDetail.RowKey}"),
                CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("RequestTypeText"), ticketDetail.RequestType),
                CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("RequestStatusText"), $"{(TicketState)ticketDetail.TicketStatus}"),
                CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("TitleDisplayText"), ticketDetail.Title),
            });
            dynamicElements.AddRange(ticketAdditionalFields);
            dynamicElements.Add(CardHelper.GetAdaptiveCardColumnSet(localizer.GetString("DescriptionText"), ticketDetail.Description));

            AdaptiveCard getTicketDetailsForSMEChatCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = dynamicElements,
                Actions = this.BuildActions(localizer),
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = getTicketDetailsForSMEChatCard,
            };
        }

        /// <summary>
        /// Return the appropriate set of card actions based on the state and information in the ticket.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Adaptive card actions.</returns>
        private List<AdaptiveAction> BuildActions(IStringLocalizer<Strings> localizer)
        {
            List<AdaptiveAction> actionsList = new List<AdaptiveAction>
            {
                this.CreateChatWithUserAction(localizer),
            };
            if (this.ticket.TicketStatus != (int)TicketState.Withdrawn)
            {
                actionsList.Add(new AdaptiveShowCardAction
                {
                    Title = localizer.GetString("ChangeStatusButtonText"),
                    Card = new AdaptiveCard(Constants.AdaptiveCardVersion)
                    {
                        Body = new List<AdaptiveElement>
                        {
                            this.GetAdaptiveChoiceSetInput(localizer),
                        },
                        Actions = new List<AdaptiveAction>
                        {
                            new AdaptiveSubmitAction
                            {
                                Title = localizer.GetString("UpdateActionText"),
                                Data = new ChangeTicketStatus
                                {
                                    TicketId = this.ticket.TicketId,
                                },
                            },
                        },
                    },
                });
                actionsList.Add(new AdaptiveShowCardAction
                {
                    Title = localizer.GetString("SeverityButtonText"),
                    Card = new AdaptiveCard(Constants.AdaptiveCardVersion)
                    {
                        Body = new List<AdaptiveElement>
                        {
                            new AdaptiveChoiceSetInput
                            {
                                Choices = new List<AdaptiveChoice>
                                {
                                    new AdaptiveChoice
                                    {
                                        Title = localizer.GetString("NormalText"),
                                        Value = Constants.NormalString,
                                    },
                                    new AdaptiveChoice
                                    {
                                        Title = localizer.GetString("UrgentText"),
                                        Value = Constants.UrgentString,
                                    },
                                },
                                Id = "RequestType",
                                Value = this.ticket.RequestType,
                                Style = AdaptiveChoiceInputStyle.Expanded,
                            },
                        },
                        Actions = new List<AdaptiveAction>
                        {
                            new AdaptiveSubmitAction
                            {
                                Title = localizer.GetString("UpdateActionText"),
                                Data = new ChangeTicketStatus
                                {
                                    TicketId = this.ticket.TicketId,
                                },
                            },
                        },
                    },
                });
            }

            return actionsList;
        }

        /// <summary>
        /// Create an adaptive card action that starts a chat with the user.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Adaptive card action for starting chat with user.</returns>
        private AdaptiveAction CreateChatWithUserAction(IStringLocalizer<Strings> localizer)
        {
            var messageToSend = localizer.GetString("SmeUserChatMessage", this.ticket.TicketId);
            var encodedMessage = Uri.EscapeDataString(messageToSend);

            return new AdaptiveOpenUrlAction
            {
                Title = localizer.GetString("ChatTextButton"),
                Url = new Uri($"https://teams.microsoft.com/l/chat/0/0?users={Uri.EscapeDataString(this.ticket.CreatedByUserPrincipalName)}&message={encodedMessage}"),
            };
        }

        /// <summary>
        /// Return the appropriate status choices based on the state and information in the ticket.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>An adaptive element which contains the drop down choices.</returns>
        private AdaptiveChoiceSetInput GetAdaptiveChoiceSetInput(IStringLocalizer<Strings> localizer)
        {
            AdaptiveChoiceSetInput choiceSet = new AdaptiveChoiceSetInput
            {
                Id = nameof(ChangeTicketStatus.Action),
                IsMultiSelect = false,
                Style = AdaptiveChoiceInputStyle.Compact,
            };

            if (this.ticket.TicketStatus == (int)TicketState.Unassigned)
            {
                choiceSet.Value = ChangeTicketStatus.AssignToSelfAction;
                choiceSet.Choices = new List<AdaptiveChoice>
                {
                    new AdaptiveChoice
                    {
                        Title = localizer.GetString("AssignToMeActionChoiceTitle"),
                        Value = ChangeTicketStatus.AssignToSelfAction,
                    },
                    new AdaptiveChoice
                    {
                        Title = localizer.GetString("CloseActionChoiceTitle"),
                        Value = ChangeTicketStatus.CloseAction,
                    },
                };
            }
            else if (this.ticket.TicketStatus == (int)TicketState.Closed)
            {
                choiceSet.Value = localizer.GetString("ReopenActionChoiceTitle");
                choiceSet.Choices = new List<AdaptiveChoice>
                {
                    new AdaptiveChoice
                    {
                        Title = localizer.GetString("ReopenActionChoiceTitle"),
                        Value = ChangeTicketStatus.ReopenAction,
                    },
                    new AdaptiveChoice
                    {
                        Title = localizer.GetString("ReopenAssignToMeActionChoiceTitle"),
                        Value = ChangeTicketStatus.AssignToSelfAction,
                    },
                };
            }
            else if (this.ticket.TicketStatus == (int)TicketState.Assigned)
            {
                choiceSet.Value = localizer.GetString("CloseActionChoiceTitle");
                choiceSet.Choices = new List<AdaptiveChoice>
                {
                    new AdaptiveChoice
                    {
                        Title = localizer.GetString("UnassignActionChoiceTitle"),
                        Value = ChangeTicketStatus.ReopenAction,
                    },
                    new AdaptiveChoice
                    {
                        Title = localizer.GetString("CloseActionChoiceTitle"),
                        Value = ChangeTicketStatus.CloseAction,
                    },
                };
            }

            return choiceSet;
        }
    }
}
