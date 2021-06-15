// <copyright file="OnCallExpertsDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Class contains details of the on call experts.
    /// </summary>
    public class OnCallExpertsDetail
    {
        /// <summary>
        /// Gets or sets list of on call experts.
        /// </summary>
        [JsonProperty("oncallexpertslist")]
        #pragma warning disable CA2227 // Collection properties should be read only - Need to set this property from json response from client Application.
        public List<string> OnCallExperts { get; set; }
        #pragma warning restore CA2227 // Collection properties should be read only

        /// <summary>
        /// Gets or sets unique identifier of the on call support created.
        /// </summary>
        [JsonProperty("oncallsupportid")]
        public string OnCallSupportId { get; set; }

        /// <summary>
        /// Gets or sets card activity id which need to be refreshed in channel with updated details.
        /// </summary>
        [JsonProperty("oncallsupportcardactivityid")]
        public string OnCallSupportCardActivityId { get; set; }
    }
}
