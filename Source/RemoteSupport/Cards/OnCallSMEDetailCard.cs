// <copyright file="OnCallSMEDetailCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Cards
{
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RemoteSupport.Common;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.Teams.Apps.RemoteSupport.Helpers;
    using Microsoft.Teams.Apps.RemoteSupport.Models;
    using Newtonsoft.Json;

    /// <summary>
    ///  Provides adaptive cards for managing on call support team details and viewing on call experts update history.
    /// </summary>
    public static class OnCallSMEDetailCard
    {
        /// <summary>
        /// Gets on call SME detail card.
        /// </summary>
        /// <param name="onCallSupportDetails"> Collection of last 10 modified on call support team details.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns an attachment of card showing on call support details.</returns>
        public static Attachment GetOnCallSMEDetailCard(IEnumerable<OnCallSupportDetail> onCallSupportDetails, IStringLocalizer<Strings> localizer)
        {
            string onCallSMENames = string.Empty;
            bool isOnCallExpertConfigured = true;

            if (onCallSupportDetails != null && onCallSupportDetails.Any())
            {
                var onCallSMEDetail = JsonConvert.DeserializeObject<List<OnCallSMEDetail>>(onCallSupportDetails.First().OnCallSMEs);
                if (onCallSMEDetail != null)
                {
                    onCallSMENames = string.Join(", ", onCallSMEDetail.Select(onCallSME => onCallSME.Name)).TrimStart(',');
                }
            }
            else
            {
                onCallSMENames = localizer.GetString("NoOnCallExpertsConfiguredText");
                isOnCallExpertConfigured = false;
            }

            AdaptiveCard adaptiveCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = isOnCallExpertConfigured ? localizer.GetString("OnCallSMEDetailCardText") : string.Empty,
                        Spacing = AdaptiveSpacing.Medium,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = onCallSMENames,
                        Weight = AdaptiveTextWeight.Bolder,
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                         Title = localizer.GetString("ManageExpertsActionText"),
                         Data = new AdaptiveCardAction
                         {
                             MsteamsCardAction = new CardAction
                             {
                                Type = Constants.FetchActionType,
                             },
                             Command = Constants.ManageExpertsAction,
                         },
                    },
                    new AdaptiveShowCardAction()
                    {
                        Title = localizer.GetString("UpdateHistoryActionText"),
                        Card = OnCallSMEUpdateHistoryCard(onCallSupportDetails, localizer),
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
            };
        }

        /// <summary>
        /// Card to show confirmation on selecting withdraw action.
        /// </summary>
        /// <param name="onCallSupportDetails"> Collection of class containing details of on call support team.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>An attachment with confirmation(yes/no)card.</returns>
        public static AdaptiveCard OnCallSMEUpdateHistoryCard(IEnumerable<OnCallSupportDetail> onCallSupportDetails, IStringLocalizer<Strings> localizer)
        {
            AdaptiveContainer container = new AdaptiveContainer();
            if (onCallSupportDetails != null && onCallSupportDetails.Any())
            {
                foreach (var onCallSupportDetail in onCallSupportDetails)
                {
                    AdaptiveColumnSet columnSet = new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = "1",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = onCallSupportDetail.ModifiedByName,
                                        Size = AdaptiveTextSize.Medium,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = "1",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = CardHelper.AdaptiveTextParseWithDateTime(onCallSupportDetail.ModifiedOn?.ToString(CultureInfo.InvariantCulture)),
                                        Size = AdaptiveTextSize.Medium,
                                    },
                                },
                            },
                        },
                    };

                    container.Items.Add(columnSet);
                }
            }
            else
            {
                container.Items.Add(new AdaptiveTextBlock
                {
                    Text = localizer.GetString("NoUpdateHistoryText"),
                    Weight = AdaptiveTextWeight.Bolder,
                    Size = AdaptiveTextSize.Medium,
                });
            }

            AdaptiveCard onCallSupportUpdateHistoryCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = "5",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = localizer.GetString("NameTitleText"),
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Size = AdaptiveTextSize.Medium,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = "5",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = localizer.GetString("DatetimeTitleText"),
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Size = AdaptiveTextSize.Medium,
                                    },
                                },
                            },
                        },
                    },

                    container,
                },
            };

            return onCallSupportUpdateHistoryCard;
        }
    }
}
