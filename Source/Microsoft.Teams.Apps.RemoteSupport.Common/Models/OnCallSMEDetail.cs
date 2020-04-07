// <copyright file="OnCallSMEDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Class contains details of on call support experts.
    /// </summary>
    public class OnCallSMEDetail
    {
        /// <summary>
        /// Gets or sets name of on call expert.
        /// </summary>
        [JsonProperty("name")]
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory object Id.
        /// </summary>
        [JsonProperty("objectid")]
        public string ObjectId { get; set; }
    }
}
