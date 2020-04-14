// <copyright file="TicketIdGenerator.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Class contains latest ticket Id details.
    /// </summary>
    public class TicketIdGenerator : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TicketIdGenerator"/> class.
        /// </summary>
        public TicketIdGenerator()
        {
            this.PartitionKey = Constants.TicketIdGeneratorPartitionKey;
        }

        /// <summary>
        /// Gets or sets ticket id.
        /// </summary>
        public int MaxTicketId { get; set; }
    }
}