// <copyright file="StorageBaseProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.RetryPolicies;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Class handles initialization to Azure table storage.
    /// </summary>
    public class StorageBaseProvider
    {
        /// <summary>
        /// Azure storage table name to perform operations.
        /// </summary>
        private readonly string tableName;

        /// <summary>
        /// Connection string of azure table storage.
        /// </summary>
        private readonly string connectionString;

        /// <summary>
        /// A lazy task to initialize Azure table storage.
        /// </summary>
        private readonly Lazy<Task> initializeTask;

        /// <summary>
        /// Azure cloud table client.
        /// </summary>
        private CloudTableClient cloudTableClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="StorageBaseProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        /// <param name="tableName">Table name of azure table storage to initialize.</param>
        public StorageBaseProvider(IOptionsMonitor<StorageOptions> storageOptions, string tableName)
        {
            storageOptions = storageOptions ?? throw new ArgumentNullException(nameof(storageOptions));
            this.connectionString = storageOptions.CurrentValue.ConnectionString;
            this.tableName = tableName;
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync());
        }

        /// <summary>
        /// Gets or sets cloud table for storing group activity and details regarding sending notification.
        /// </summary>
        protected CloudTable CloudTable { get; set; }

        /// <summary>
        /// Ensures Microsoft Azure Table Storage should be created before working on table.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        protected async Task EnsureInitializedAsync()
        {
            await this.initializeTask.Value;
        }

        /// <summary>
        /// Create storage table if it does not exist.
        /// </summary>
        /// <returns><see cref="Task"/> representing the asynchronous operation task which represents table is created if it does not exists.</returns>
        private async Task<CloudTable> InitializeAsync()
        {
            // Exponential retry policy with back off set to 3 seconds and 5 retries.
            var exponentialRetryPolicy = new TableRequestOptions()
            {
                RetryPolicy = new ExponentialRetry(TimeSpan.FromSeconds(3), 5),
            };

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(this.connectionString);
            this.cloudTableClient = storageAccount.CreateCloudTableClient();
            this.cloudTableClient.DefaultRequestOptions = exponentialRetryPolicy;
            this.CloudTable = this.cloudTableClient.GetTableReference(this.tableName);

            await this.CloudTable.CreateIfNotExistsAsync();

            return this.CloudTable;
        }
    }
}