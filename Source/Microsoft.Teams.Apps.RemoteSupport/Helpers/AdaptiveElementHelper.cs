// <copyright file="AdaptiveElementHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using AdaptiveCards;
    using Microsoft.Teams.Apps.RemoteSupport.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// Helper class to convert Json property into Adaptive card element
    /// </summary>
    public static class AdaptiveElementHelper
    {
        /// <summary>
        /// Converts json property to adaptive card TextBlock element.
        /// </summary>
        /// <param name="jsonProperty">TextBlock item element json property.</param>
        /// <param name="showDateValidation">true if need to show validation message else false.</param>
        /// <returns>Returns adaptive card TextBlock item element.</returns>
        public static AdaptiveTextBlock ConvertToAdaptiveTextBlock(string jsonProperty, bool showDateValidation = false)
        {
            var result = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonProperty);
            bool isVisible = true;
            if (!string.IsNullOrEmpty(CardHelper.TryParseTicketDetailsKeyValuePair(result, "isVisible")))
            {
                bool status = bool.TryParse(CardHelper.TryParseTicketDetailsKeyValuePair(result, "isVisible"), out isVisible);
            }

            string color = CardHelper.TryParseTicketDetailsKeyValuePair(result, "color");
            return new AdaptiveTextBlock()
            {
                Id = CardHelper.TryParseTicketDetailsKeyValuePair(result, "id"),
                Text = CardHelper.TryParseTicketDetailsKeyValuePair(result, "text"),
                IsVisible = isVisible,
                Color = string.IsNullOrEmpty(color) ? AdaptiveTextColor.Default : (AdaptiveTextColor)Enum.Parse(typeof(AdaptiveTextColor), color),
            };
        }

        /// <summary>
        /// Converts json property to adaptive card TextInput element.
        /// </summary>
        /// <param name="jsonProperty">TextInput item element json property.</param>
        /// <returns>Returns adaptive card TextInput item element.</returns>
        public static AdaptiveTextInput ConvertToAdaptiveTextInput(string jsonProperty)
        {
            var result = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonProperty);
            int maxLength = 500;
            bool flag = int.TryParse(CardHelper.TryParseTicketDetailsKeyValuePair(result, "maxLength"), out maxLength);

            return new AdaptiveTextInput()
            {
                Id = CardHelper.TryParseTicketDetailsKeyValuePair(result, "id"),
                Placeholder = CardHelper.TryParseTicketDetailsKeyValuePair(result, "placeholder"),
                Value = CardHelper.TryParseTicketDetailsKeyValuePair(result, "value"),
                MaxLength = maxLength,
            };
        }

        /// <summary>
        /// Converts json property to adaptive card DateInput element.
        /// </summary>
        /// <param name="jsonProperty">DateInput item element json property.</param>
        /// <returns>Returns adaptive card DateInput item element.</returns>
        public static AdaptiveDateInput ConvertToAdaptiveDateInput(string jsonProperty)
        {
            var result = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonProperty);

            return new AdaptiveDateInput()
            {
                Id = CardHelper.TryParseTicketDetailsKeyValuePair(result, "id"),
                Placeholder = CardHelper.TryParseTicketDetailsKeyValuePair(result, "placeholder"),
                Value = string.IsNullOrEmpty(CardHelper.TryParseTicketDetailsKeyValuePair(result, "value")) ? DateTime.Now.ToString(CultureInfo.InvariantCulture) : CardHelper.TryParseTicketDetailsKeyValuePair(result, "value"),
                Max = CardHelper.TryParseTicketDetailsKeyValuePair(result, "max"),
                Min = CardHelper.TryParseTicketDetailsKeyValuePair(result, "min"),
            };
        }

        /// <summary>
        /// Converts JSON property to adaptive card ChoiceSetInput element.
        /// </summary>
        /// <param name="jsonProperty">ChoiceSetInput item element json property.</param>
        /// <returns>Returns adaptive card ChoiceSetInput item element.</returns>
        public static AdaptiveChoiceSetInput ConvertToAdaptiveChoiceSetInput(string jsonProperty)
        {
            var adpativeChoiceSetCard = JsonConvert.DeserializeObject<InputChoiceSet>(jsonProperty);
            List<AdaptiveChoice> choices = adpativeChoiceSetCard.Choices
                .Select(choice => new AdaptiveChoice()
                {
                    Title = choice.Title,
                    Value = choice.Value,
                })
                .ToList();

            return new AdaptiveChoiceSetInput()
            {
                IsMultiSelect = adpativeChoiceSetCard.IsMultiSelect,
                Choices = choices,
                Id = adpativeChoiceSetCard.Id,
                Style = adpativeChoiceSetCard.Style,
                Value = adpativeChoiceSetCard.Value,
            };
        }
    }
}
