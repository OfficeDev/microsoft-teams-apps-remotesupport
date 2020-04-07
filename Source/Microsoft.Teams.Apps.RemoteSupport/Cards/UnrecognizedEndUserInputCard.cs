// <copyright file="UnrecognizedEndUserInputCard.cs" company="Microsoft">
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
    ///  This class handles unrecognized input sent by the team member-sending random text to bot.
    /// </summary>
    public static class UnrecognizedEndUserInputCard
    {
        /// <summary>
        /// This method will construct adaptive card when unrecognized input sent by user.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Card attachment.</returns>
        public static Attachment GetCard(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard unRecognisedCard = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("EndUserCustomMessage"),
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
                                Text = Constants.NewRequestAction,
                            },
                        },
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = unRecognisedCard,
            };
        }
    }
}