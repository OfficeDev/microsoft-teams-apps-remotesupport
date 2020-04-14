// <copyright file="TicketDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Azure.Search;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Class contains details of tickets created in table storage.
    /// </summary>
    public class TicketDetail : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TicketDetail"/> class.
        /// Constructor method used to initialize partition key of table.
        /// </summary>
        public TicketDetail()
        {
            this.PartitionKey = Constants.TicketDetailPartitionKey;
        }

        /// <summary>
        /// Gets or sets unique identifier of the ticket created.
        /// </summary>
        [Key]
        [IsSearchable]
        [JsonProperty("TicketId")]
        public string TicketId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets the display name of the assigned SME currently working on the ticket.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("AssignedToName")]
        public string AssignedToName { get; set; }

        /// <summary>
        /// Gets or sets the AAD object id of the assigned SME currently working on the ticket.
        /// </summary>
        [JsonProperty("AssignedToObjectId")]
        public string AssignedToObjectId { get; set; }

        /// <summary>
        /// Gets or sets the ticket title.
        /// </summary>
        [IsSearchable]
        [JsonProperty("Title")]
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the ticket description.
        /// </summary>
        [IsSearchable]
        [JsonProperty("Description")]
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets the  UTC date and time the ticket was closed.
        /// </summary>
        [IsSortable]
        [JsonProperty("ClosedOn")]
        public string ClosedOn { get; set; }

        /// <summary>
        /// Gets or sets the display name of the user who closed the ticket.
        /// </summary>
        [JsonProperty("ClosedByName")]
        public string ClosedByName { get; set; }

        /// <summary>
        /// Gets or sets the display name of the user who last modified the ticket.
        /// </summary>
        [JsonProperty("LastModifiedByName")]
        public string LastModifiedByName { get; set; }

        /// <summary>
        /// Gets or sets the  UTC date and time the ticket was last modified.
        /// </summary>
        [IsSortable]
        [JsonProperty("LastModifiedOn")]
        public DateTimeOffset? LastModifiedOn { get; set; }

        /// <summary>
        /// Gets or sets the date and time on which the request was first occurred on.
        /// </summary>
        [IsSortable]
        [JsonProperty("IssueOccuredOn")]
        public virtual DateTimeOffset IssueOccuredOn { get; set; }

        /// <summary>
        /// Gets or sets the AAD object id of the user that last modified the ticket.
        /// </summary>
        [JsonProperty("LastModifiedByObjectId")]
        public string LastModifiedByObjectId { get; set; }

        /// <summary>
        /// Gets or sets the AAD object id of the user who closed the ticket.
        /// </summary>
        [JsonProperty("ClosedByObjectId")]
        public string ClosedByObjectId { get; set; }

        /// <summary>
        /// Gets time stamp from storage table.
        /// </summary>
        [IsSortable]
        [JsonProperty("Timestamp")]
        public new DateTimeOffset Timestamp => base.Timestamp;

        /// <summary>
        /// Gets or sets the display name of the user who created the ticket.
        /// </summary>
        [IsSearchable]
        [JsonProperty("RequesterName")]
        public string RequesterName { get; set; }

        /// <summary>
        /// Gets or sets the user principal name (UPN) of the user that created the ticket.
        /// </summary>
        [JsonProperty("CreatedByUserPrincipalName")]
        public string CreatedByUserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets the conversation id of the 1:1 chat with the user that created the ticket.
        /// </summary>
        [JsonProperty("RequesterConversationId")]
        public string RequesterConversationId { get; set; }

        /// <summary>
        /// Gets or sets the activity id of the root card in the personal.
        /// </summary>
        [JsonProperty("RequesterTicketActivityId")]
        public string RequesterTicketActivityId { get; set; }

        /// <summary>
        /// Gets or sets the activity id of the root card in the SME channel.
        /// </summary>
        [JsonProperty("SmeTicketActivityId")]
        public string SmeTicketActivityId { get; set; }

        /// <summary>
        /// Gets or sets the severity of the ticketId.
        /// </summary>
        [JsonProperty("RequestType")]
        public string RequestType { get; set; }

        /// <summary>
        /// Gets or sets the conversation id of the thread pertaining to this ticket in the SME channel.
        /// </summary>
        [JsonProperty("SmeConversationId")]
        public string SmeConversationId { get; set; }

        /// <summary>
        /// Gets or sets status of the ticket.
        /// </summary>
        [IsSortable]
        [IsFilterable]
        [JsonProperty("TicketStatus")]
        public int? TicketStatus { get; set; }

        /// <summary>
        /// Gets or sets status of the ticket.
        /// </summary>
        [IsSortable]
        [IsFilterable]
        [JsonProperty("Severity")]
        public int? Severity { get; set; }

        /// <summary>
        /// Gets or sets the date and time on when the ticket is created on.
        /// </summary>
        [JsonProperty("CreatedOn")]
        public DateTimeOffset CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory objectId of user who created ticket.
        /// </summary>
        [IsSortable]
        [IsFilterable]
        [JsonProperty("CreatedByObjectId")]
        public string CreatedByObjectId { get; set; }

        /// <summary>
        /// Gets or sets the ticket additional properties.
        /// </summary>
        [JsonProperty("AdditionalProperties")]
        public string AdditionalProperties { get; set; }

        /// <summary>
        /// Gets or sets the new card id.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("CardId")]
        public string CardId { get; set; }
    }
}
