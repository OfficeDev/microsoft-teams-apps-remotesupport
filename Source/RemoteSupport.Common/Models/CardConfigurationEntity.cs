// <copyright file="CardConfigurationEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    using System;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Serialization;

    /// <summary>
    /// Class contains details of card configuration created in table storage.
    /// </summary>
    public class CardConfigurationEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CardConfigurationEntity"/> class.
        /// </summary>
        public CardConfigurationEntity()
        {
            this.PartitionKey = Constants.CardConfigurationPartitionKey;
        }

        /// <summary>
        /// Gets or sets GUID to uniquely identifies the Card.
        /// </summary>
        public string CardId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets Id of the team in which card configuration created.
        /// </summary>
        [JsonProperty("TeamId")]
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets url of the expert team.
        /// </summary>
        [JsonProperty("TeamLink")]
        public string TeamLink { get; set; }

        /// <summary>
        /// Gets or sets card creation time.
        /// </summary>
        public DateTime CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets the user principal name (UPN) of the user that created the ticket.
        /// </summary>
        public string CreatedByUserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory objectId of user who created ticket.
        /// </summary>
        public string CreatedByObjectId { get; set; }

        /// <summary>
        /// Gets or sets adaptive card items json properties.
        /// </summary>
        [JsonProperty("CardTemplate")]
        public string CardTemplate { get; set; }
    }
}
