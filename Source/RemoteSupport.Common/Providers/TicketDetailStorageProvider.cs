// <copyright file="TicketDetailStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Ticket provider helps in fetching and storing information in storage table.
    /// </summary>
    public class TicketDetailStorageProvider : StorageBaseProvider, ITicketDetailStorageProvider
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TicketDetailStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        public TicketDetailStorageProvider(IOptionsMonitor<StorageOptions> storageOptions)
            : base(storageOptions, Constants.TicketDetailTable)
        {
        }

        /// <summary>
        /// Store or update ticket entity in table storage.
        /// </summary>
        /// <param name="ticketDetails">Represents ticket entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents configuration entity is saved or updated.</returns>
        public async Task<bool> UpsertTicketAsync(TicketDetail ticketDetails)
        {
            var result = await this.StoreOrUpdateTicketEntityAsync(ticketDetails);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get already saved entity detail from storage table.
        /// </summary>
        /// <param name="ticketId">ticket id received from bot based on which appropriate row data will be fetched.</param>
        /// <returns><see cref="Task"/> Already saved entity detail.</returns>
        public async Task<TicketDetail> GetTicketAsync(string ticketId)
        {
            await this.EnsureInitializedAsync(); // When there is no ticket created by end user and messaging extension is open by SME, table initialization is required before creating search index or data source or indexer.
            if (string.IsNullOrEmpty(ticketId))
            {
                return null;
            }

            var searchOperation = TableOperation.Retrieve<TicketDetail>(Constants.TicketDetailPartitionKey, ticketId);
            var searchResult = await this.CloudTable.ExecuteAsync(searchOperation);

            return (TicketDetail)searchResult.Result;
        }

        /// <summary>
        /// Store or update ticket entity in table storage.
        /// </summary>
        /// <param name="ticketDetails">Represents ticket entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents ticket entity is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdateTicketEntityAsync(TicketDetail ticketDetails)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(ticketDetails);
            var result = await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
            return result;
        }
    }
}
