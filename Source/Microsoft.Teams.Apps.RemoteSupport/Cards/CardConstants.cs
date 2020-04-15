// <copyright file="CardConstants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Cards
{
    /// <summary>
    /// Constants used in bot and task module cards
    /// </summary>
    public static class CardConstants
    {
        /// <summary>
        /// Text block card id for first observed on text in remote support request card
        /// </summary>
        public const string IssueOccurredOnId = "IssueOccurredOn";

        /// <summary>
        /// Text block card id for date time validation message in remote support request card
        /// </summary>
        public const string DateValidationMessageId = "DateValidationMessage";

        /// <summary>
        /// Date time format to support adaptive card text feature.
        /// </summary>
        /// <remarks>
        /// refer adaptive card text feature https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/text-features#datetime-formatting-and-localization.
        /// </remarks>
        public const string Rfc3339DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'";
    }
}
