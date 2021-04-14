// <copyright file="WithdrawCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RemoteSupport.Common;
    using Microsoft.Teams.Apps.RemoteSupport.Models;

    /// <summary>
    /// Class holds card with confirmation card and withdraw details.
    /// </summary>
    public static class WithdrawCard
    {
        /// <summary>
        /// Get the withdraw card with new request button.
        /// </summary>
        /// <param name="requestNumber">ticketId of the user.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns an attachment withdraw card.</returns>
        public static Attachment GetCard(string requestNumber, IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard welcomeCard = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("WithdrawText", requestNumber),
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("NewRequestButtonText"),
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = Constants.MessageBackActionType,
                                Text = localizer.GetString("NewRequestButtonText"),
                            },
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = welcomeCard,
            };
        }

        /// <summary>
        /// Card to show confirmation on selecting withdraw action.
        /// </summary>
        /// <param name="ticketId">TicketId of particular request.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>An attachment with confirmation(yes/no)card.</returns>
        public static AdaptiveCard ConfirmationCard(string ticketId, IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion));
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("WithdrawConfirmation"),
                        Wrap = true,
                    },
                },
            };
            card.Body.Add(container);

            card.Actions.Add(
                new AdaptiveSubmitAction()
                {
                    Title = localizer.GetString("Yes"),
                    Data = new AdaptiveCardAction
                    {
                        MsteamsCardAction = new CardAction
                        {
                            Type = ActionTypes.MessageBack,
                            Text = Constants.WithdrawRequestAction,
                        },
                        PostedValues = ticketId,
                    },
                });

            card.Actions.Add(
                new AdaptiveSubmitAction()
                {
                    Title = localizer.GetString("No"),
                    Data = new AdaptiveCardAction
                    {
                        MsteamsCardAction = new CardAction
                        {
                            Type = ActionTypes.MessageBack,
                            Text = Constants.NoCommand,
                        },
                    },
                });

            return card;
        }
    }
}
