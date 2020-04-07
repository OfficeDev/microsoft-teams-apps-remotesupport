// <copyright file="ChangeTicketStatus.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Represents the data payload of Action.Submit to change the status of a ticket.
    /// </summary>
    public class ChangeTicketStatus
    {
        /// <summary>
        /// Action that reopens a closed ticket.
        /// </summary>
        public const string ReopenAction = "Reopen";

        /// <summary>
        /// Action that closes a ticket.
        /// </summary>
        public const string CloseAction = "Close";

        /// <summary>
        /// Action that relates to change in severity of an ticket.
        /// </summary>
        public const string RequestTypeAction = "RequestType";

        /// <summary>
        /// Action that assigns a ticket to the person that performed the action.
        /// </summary>
        public const string AssignToSelfAction = "AssignToSelf";

        /// <summary>
        /// Gets or sets the ticket id.
        /// </summary>
        [JsonProperty("ticketId")]
        public string TicketId { get; set; }

        /// <summary>
        /// Gets or sets the action to perform on the ticket.
        /// </summary>
        [JsonProperty("action")]
        public string Action { get; set; }

        /// <summary>
        /// Gets or sets the severity of the ticketId.
        /// </summary>
        [JsonProperty("RequestType")]
        public string RequestType { get; set; }
    }
}