// <copyright file="OnCallSupportDetailStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// On call support detail provider helps in fetching and storing information in storage table.
    /// </summary>
    public class OnCallSupportDetailStorageProvider : StorageBaseProvider, IOnCallSupportDetailStorageProvider
    {
        private readonly ILogger<OnCallSupportDetailStorageProvider> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="OnCallSupportDetailStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public OnCallSupportDetailStorageProvider(IOptionsMonitor<StorageOptions> storageOptions, ILogger<OnCallSupportDetailStorageProvider> logger)
        : base(storageOptions, Constants.OnCallSupportDetailTable)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Save on call support details in Azure Table Storage.
        /// </summary>
        /// <param name="onCallSupportTeamDetails">On call support details to be stored in table storage.</param>
        /// <returns><see cref="Task"/> Returns OnCallSupportId when on call support data was saved successfully.</returns>
        public async Task<string> UpsertOnCallSupportDetailsAsync(OnCallSupportDetail onCallSupportTeamDetails)
        {
            await this.EnsureInitializedAsync();
            onCallSupportTeamDetails = onCallSupportTeamDetails ?? throw new ArgumentNullException(nameof(onCallSupportTeamDetails));
            onCallSupportTeamDetails.RowKey = Guid.NewGuid().ToString();
            TableOperation addOperation = TableOperation.Insert(onCallSupportTeamDetails);
            var result = await this.CloudTable.ExecuteAsync(addOperation);
            if (result.Result != null)
            {
                var onCallSMEDetails = (OnCallSupportDetail)result.Result;
                return onCallSMEDetails.OnCallSupportId;
            }

            return string.Empty;
        }

        /// <summary>
        /// Get already saved entity detail from storage table.
        /// </summary>
        /// <param name="onCallSMEId">onCallSMEId received from bot based on which appropriate row data will be fetched.</param>
        /// <returns><see cref="Task"/> Already saved entity detail.</returns>
        public async Task<OnCallSupportDetail> GetOnCallSupportDetailAsync(string onCallSMEId)
        {
            await this.EnsureInitializedAsync(); // When there is no on call support details added and task module is opened by SME, table initialization is required before creating search index or data source or indexer.
            if (string.IsNullOrEmpty(onCallSMEId))
            {
                this.logger.LogInformation("There are no on call experts configured.");
                return null;
            }

            var searchOperation = TableOperation.Retrieve<OnCallSupportDetail>(Constants.OnCallSupportDetailPartitionKey, onCallSMEId);
            var searchResult = await this.CloudTable.ExecuteAsync(searchOperation);

            return (OnCallSupportDetail)searchResult.Result;
        }
    }
}
