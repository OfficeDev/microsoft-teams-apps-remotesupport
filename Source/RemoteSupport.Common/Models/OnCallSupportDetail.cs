// <copyright file="OnCallSupportDetail.cs" company="Microsoft">
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
    /// Class contains details of on call support team.
    /// </summary>
    public class OnCallSupportDetail : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="OnCallSupportDetail"/> class.
        /// Constructor method used to initialize partition key of table.
        /// </summary>
        public OnCallSupportDetail()
        {
            this.PartitionKey = Constants.OnCallSupportDetailPartitionKey;
        }

        /// <summary>
        /// Gets or sets unique identifier of the on call support created.
        /// </summary>
        [Key]
        [JsonProperty("OnCallSupportId")]
        public string OnCallSupportId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets time stamp from storage table.
        /// </summary>
        [IsSortable]
        [JsonProperty("Timestamp")]
        public new DateTimeOffset Timestamp => base.Timestamp;

        /// <summary>
        /// Gets or sets name of user who modified on call support experts list.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("ModifiedByName")]
        public string ModifiedByName { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of user who modified on call support experts list.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("ModifiedByObjectId")]
        public string ModifiedByObjectId { get; set; }

        /// <summary>
        /// Gets or sets date on which on call support experts list was updated.
        /// </summary>
        [IsSortable]
        [IsFilterable]
        [JsonProperty("ModifiedOn")]
        public DateTime? ModifiedOn { get; set; }

        /// <summary>
        /// Gets or sets on call support experts details in json string.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("OnCallSMEs")]
        public string OnCallSMEs { get; set; }
    }
}
