// <copyright file="TicketSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Search;
    using Microsoft.Azure.Search.Models;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;

    /// <summary>
    /// SearchService which will help in creating index, indexer and data source if it doesn't exist
    /// for indexing table which will be used for search by message extension.
    /// </summary>
    public class TicketSearchService : ITicketSearchService, IDisposable
    {
        private const string TicketsIndexName = "ticketdetaildata-index";
        private const string TicketsIndexerName = "ticketdetaildata-indexer";
        private const string TicketsDataSourceName = "ticketdetaildata-storage";

        // Default to 25 results, same as page size of a messaging extension query
        private const int DefaultSearchResultCount = 25;
        private readonly Lazy<Task> initializeTask;
        private readonly SearchServiceClient searchServiceClient;
        private readonly SearchIndexClient searchIndexClient;
        private readonly ITicketDetailStorageProvider ticketDetailStorageProvider;
        private readonly int searchIndexingIntervalInMinutes;
        private readonly ILogger<TicketSearchService> logger;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly SearchServiceOptions searchServiceOptions;

        // Flag: Has Dispose already been called?
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="TicketSearchService"/> class.
        /// </summary>
        /// <param name="searchServiceOptions">A set of key/value application configuration properties.</param>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        /// <param name="ticketDetailStorageProvider"> TicketsProvider provided by dependency injection.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public TicketSearchService(
            IOptionsMonitor<SearchServiceOptions> searchServiceOptions,
            IOptionsMonitor<StorageOptions> storageOptions,
            ITicketDetailStorageProvider ticketDetailStorageProvider,
            ILogger<TicketSearchService> logger)
        {
            searchServiceOptions = searchServiceOptions ?? throw new ArgumentNullException(nameof(searchServiceOptions));
            storageOptions = storageOptions ?? throw new ArgumentNullException(nameof(storageOptions));

            this.searchServiceOptions = searchServiceOptions.CurrentValue;
            string searchServiceValue = this.searchServiceOptions.SearchServiceName;
            this.searchServiceClient = new SearchServiceClient(
                searchServiceValue,
                new SearchCredentials(this.searchServiceOptions.SearchServiceAdminApiKey));
            this.searchIndexClient = new SearchIndexClient(
                searchServiceValue,
                TicketsIndexName,
                new SearchCredentials(this.searchServiceOptions.SearchServiceQueryApiKey));
            this.searchIndexingIntervalInMinutes = Convert.ToInt32(this.searchServiceOptions.SearchIndexingIntervalInMinutes, CultureInfo.InvariantCulture);

            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(storageOptions.CurrentValue.ConnectionString));
            this.ticketDetailStorageProvider = ticketDetailStorageProvider;
            this.logger = logger;
        }

        /// <summary>
        /// Provide search result for table to be used by SME based on Azure search service.
        /// </summary>
        /// <param name="searchScope">Scope of the search.</param>
        /// <param name="searchQuery">searchQuery to be provided by message extension.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="requestorId">Requester id of the user to get specific tickets.</param>
        /// <returns>List of search results.</returns>
        public async Task<IList<TicketDetail>> SearchTicketsAsync(TicketSearchScope searchScope, string searchQuery, int? count = null, int? skip = null, string requestorId = "")
        {
            await this.EnsureInitializedAsync();

            IList<TicketDetail> tickets = new List<TicketDetail>();

            SearchParameters searchParameters = new SearchParameters();
            switch (searchScope)
            {
                case TicketSearchScope.UrgentTickets:
                    searchParameters.Filter = $"Severity eq {(int)TicketSeverity.Urgent}";
                    searchParameters.OrderBy = new[] { "Timestamp desc" };
                    break;

                case TicketSearchScope.AssignedTickets:
                    searchParameters.Filter = $"TicketStatus eq {(int)TicketState.Assigned}";
                    searchParameters.OrderBy = new[] { "Timestamp desc" };
                    break;

                case TicketSearchScope.UnassignedTickets:
                    searchParameters.Filter = $"TicketStatus eq {(int)TicketState.Unassigned}";
                    searchParameters.OrderBy = new[] { "Timestamp desc" };
                    break;
                case TicketSearchScope.ActiveTickets:
                    searchParameters.Filter = $"(TicketStatus eq {(int)TicketState.Assigned} or TicketStatus eq {(int)TicketState.Unassigned}) and CreatedByObjectId eq '{requestorId}'";
                    searchParameters.OrderBy = new[] { "Timestamp desc" };
                    break;
                case TicketSearchScope.ClosedTickets:
                    searchParameters.Filter = $"TicketStatus eq {(int)TicketState.Closed} and CreatedByObjectId eq '{requestorId}'";
                    searchParameters.OrderBy = new[] { "Timestamp desc" };
                    break;
            }

            searchParameters.Top = count ?? DefaultSearchResultCount;
            searchParameters.Skip = skip ?? 0;
            searchParameters.IncludeTotalResultCount = false;
            searchParameters.Select = new[] { "Title", "TicketStatus", "AssignedToName", "AssignedToObjectId", "CreatedOn", "RequesterName", "CreatedByUserPrincipalName", "Description", "RequesterName", "SmeConversationId", "SmeTicketActivityId", "ClosedOn", "ClosedByName", "LastModifiedByName", "Severity", "RequesterConversationId", "RequesterTicketActivityId", "TicketId", "RequestType", "CreatedByObjectId" };

            var docs = await this.searchIndexClient.Documents.SearchAsync<TicketDetail>(searchQuery, searchParameters);

            if (docs != null)
            {
                tickets = docs.Results.Select(result => result.Document).ToList();
            }

            this.logger.LogInformation("Retrieved documents from ticket search service successfully.");
            return tickets;
        }

        /// <summary>
        /// This code added to correctly implement the disposable pattern.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">True if already disposed else false.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.searchIndexClient.Dispose();
                this.searchServiceClient.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Create index, indexer and data source it doesn't exist.
        /// </summary>
        /// <param name="storageConnectionString">Connection string to the data store.</param>
        /// <returns>Tracking task.</returns>
        private async Task InitializeAsync(string storageConnectionString)
        {
            await this.ticketDetailStorageProvider.GetTicketAsync(string.Empty); // When there is no ticket created by end user and messaging extension is open by SME, table initialization is required here before creating search index or data source or indexer.
            await this.CreateIndexAsync();
            await this.CreateDataSourceAsync(storageConnectionString);
            await this.CreateIndexerAsync();
        }

        /// <summary>
        /// Create index in Azure search service if it doesn't exist.
        /// </summary>
        /// <returns><see cref="Task"/> That represents index is created if it is not created.</returns>
        private async Task CreateIndexAsync()
        {
            if (!this.searchServiceClient.Indexes.Exists(TicketsIndexName))
            {
                var tableIndex = new Index()
                {
                    Name = TicketsIndexName,
                    Fields = FieldBuilder.BuildForType<TicketDetail>(),
                };
                await this.searchServiceClient.Indexes.CreateAsync(tableIndex);
            }
        }

        /// <summary>
        /// Add data source if it doesn't exist in Azure search service.
        /// </summary>
        /// <param name="connectionString">Connection string to the data store.</param>
        /// <returns><see cref="Task"/> That represents data source is added to Azure search service.</returns>
        private async Task CreateDataSourceAsync(string connectionString)
        {
            if (!this.searchServiceClient.DataSources.Exists(TicketsDataSourceName))
            {
                var dataSource = DataSource.AzureTableStorage(
                    name: TicketsDataSourceName,
                    storageConnectionString: connectionString,
                    tableName: Constants.TicketDetailTable);

                await this.searchServiceClient.DataSources.CreateAsync(dataSource);
            }
        }

        /// <summary>
        /// Create indexer if it doesn't exist in Azure search service.
        /// </summary>
        /// <returns><see cref="Task"/> That represents indexer is created if not available in Azure search service.</returns>
        private async Task CreateIndexerAsync()
        {
                if (!this.searchServiceClient.Indexers.Exists(TicketsIndexerName))
                {
                    var indexer = new Indexer()
                    {
                        Name = TicketsIndexerName,
                        DataSourceName = TicketsDataSourceName,
                        TargetIndexName = TicketsIndexName,
                        Schedule = new IndexingSchedule(TimeSpan.FromMinutes(this.searchIndexingIntervalInMinutes)),
                    };

                    await this.searchServiceClient.Indexers.CreateAsync(indexer);
                }
        }

        /// <summary>
        /// Initialization of InitializeAsync method which will help in indexing.
        /// </summary>
        /// <returns>Task with initialized data.</returns>
        private Task EnsureInitializedAsync()
        {
            return this.initializeTask.Value;
        }
    }
}
