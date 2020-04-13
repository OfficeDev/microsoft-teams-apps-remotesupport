// <copyright file="AdaptiveCardPlaceHolderMapper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Class maps controls from json card file to adaptive card.
    /// </summary>
    public class AdaptiveCardPlaceHolderMapper
    {
        /// <summary>
        /// Gets or sets unique identifier of the control.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets type of control to be shown on adaptive card.
        /// </summary>
        [JsonProperty("type")]
        public string InputType { get; set; }

        /// <summary>
        /// Gets or sets displayName of control to be shown on adaptive card.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }
    }
}
