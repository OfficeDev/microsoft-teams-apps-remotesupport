// <copyright file="InputChoiceSet.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Models
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Newtonsoft.Json;

    /// <summary>
    /// InputChoiceSet
    /// </summary>
    public class InputChoiceSet
    {
        /// <summary>
        /// Gets or Sets type
        /// </summary>
        [JsonProperty("type")]
        public string Type { get; set; }

        /// <summary>
        /// Gets Choices.
        /// </summary>
        [JsonProperty("choices")]
        public List<AdaptiveChoiceSet> Choices { get; } = new List<AdaptiveChoiceSet>();

        /// <summary>
        /// Gets or sets a value indicating whether gets or Sets indicating whether gets or sets check box is enabled or not.
        /// </summary>
        [JsonProperty("isMultiSelect")]
        public bool IsMultiSelect { get; set; }

        /// <summary>
        /// Gets or Sets Input id.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets style.
        /// </summary>
        public AdaptiveChoiceInputStyle Style { get; set; }

        /// <summary>
        /// Gets or sets input elemenet value.
        /// </summary>
        public string Value { get; set; }
    }
}
