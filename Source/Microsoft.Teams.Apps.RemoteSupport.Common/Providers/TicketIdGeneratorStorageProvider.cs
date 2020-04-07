// <copyright file="TicketIdGeneratorStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Class generate ticket ids from table storage.
    /// </summary>
    public class TicketIdGeneratorStorageProvider : StorageBaseProvider, ITicketIdGeneratorStorageProvider
    {
        /// <summary>
        /// Table name which stores Ticket id for the new request ticket.
        /// </summary>
        public const string TicketIdGeneratorTableName = "TicketIdGenerator";

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<TicketIdGeneratorStorageProvider> logger;

        /// <summary>
        /// Represents retry attempt count for '412 - Precondition Failed' exception.
        /// </summary>
        private int retryCount;

        /// <summary>
        /// Initializes a new instance of the <see cref="TicketIdGeneratorStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public TicketIdGeneratorStorageProvider(IOptionsMonitor<StorageOptions> storageOptions, ILogger<TicketIdGeneratorStorageProvider> logger)
        : base(storageOptions, TicketIdGeneratorTableName)
        {
            this.logger = logger;
            this.retryCount = 0;
        }

        /// <summary>
        /// Gets the max ticket id from the table for new request created.
        /// </summary>
        /// <returns>Ticket Id.</returns>
        public async Task<int> GetTicketIdAsync()
        {
            int nextTicketId = 0;
            try
            {
                await this.EnsureInitializedAsync();
                TableQuery<TicketIdGenerator> query = new TableQuery<TicketIdGenerator>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, Constants.TicketIdGeneratorPartitionKey));
                TableContinuationToken tableContinuationToken = null;
                do
                {
                    var queryResponse = await this.CloudTable.ExecuteQuerySegmentedAsync(query, tableContinuationToken);
                    tableContinuationToken = queryResponse.ContinuationToken;
                    var ticketIdGenerator = queryResponse.Results.FirstOrDefault();
                    if (ticketIdGenerator == null)
                    {
                        ticketIdGenerator = new TicketIdGenerator
                        {
                            MaxTicketId = 1,
                            RowKey = Guid.NewGuid().ToString(),
                        };
                        TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(ticketIdGenerator);
                        TableResult result = await this.CloudTable.ExecuteAsync(insertOrMergeOperation);
                        nextTicketId = ticketIdGenerator.MaxTicketId;
                    }
                    else
                    {
                        await this.UpdateTicketIdAsync(ticketIdGenerator);
                        nextTicketId = ticketIdGenerator.MaxTicketId;
                        this.retryCount = 0;
                    }
                }
                while (tableContinuationToken != null && query != null);
            }
            catch (StorageException ex)
            {
                if (ex.RequestInformation.HttpStatusCode == 412)
                {
                    this.logger.LogError("Optimistic concurrency violation – entity has changed since it was retrieved.");
                    await this.RetryTicketIdGenerationAsync();
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError($"Error {ex.Message}");
                throw new Exception(ex.Message);
            }

            return nextTicketId;
        }

        /// <summary>
        /// update the ticket id in the table storage.
        /// </summary>
        /// <param name="ticketIdGenerator"> Entity containing latest ticket Id details.</param>
        /// <returns> Returns next ticket id generated from table storage.</returns>
        public async Task<int> UpdateTicketIdAsync(TicketIdGenerator ticketIdGenerator)
        {
            if (ticketIdGenerator != null)
            {
                ticketIdGenerator.MaxTicketId += 1;
                TableOperation replaceOperation = TableOperation.Replace(ticketIdGenerator);
                await this.CloudTable.ExecuteAsync(replaceOperation);
                return ticketIdGenerator.MaxTicketId;
            }

            return 0;
        }

        /// <summary>
        /// Retries ticket Id generation in case of '412 - Precondition Failed' exception.
        /// </summary>
        /// <returns> Returns ticket Id generated in case of success and throws error in case of max retry attempt.</returns>
        private async Task<int> RetryTicketIdGenerationAsync()
        {
            // Retry for getting latest updated ticket Id from table in case other user has updated value in table.
            if (this.retryCount < 3)
            {
                this.retryCount++;
                return await this.GetTicketIdAsync();
            }

            this.retryCount = 0;
            throw new Exception("Retry limit exceeded for precondition failed exception.");
        }
    }
}
