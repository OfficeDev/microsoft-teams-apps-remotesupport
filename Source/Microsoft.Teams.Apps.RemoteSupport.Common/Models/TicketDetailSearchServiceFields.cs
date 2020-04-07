// <copyright file="TicketDetailSearchServiceFields.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    using System;
    using Microsoft.Azure.Search;
    using Newtonsoft.Json;

    /// <summary>
    /// Class contains fields of tickets created in table storage.
    /// </summary>
    public class TicketDetailSearchServiceFields : TicketDetail
    {
        /// <summary>
        /// Gets or sets the date and time on which the request was first occurred on.
        /// </summary>
        [IsSortable]
        [JsonProperty("IssueOccuredOn")]
        public override DateTimeOffset IssueOccuredOn { get; set; }
    }
}
