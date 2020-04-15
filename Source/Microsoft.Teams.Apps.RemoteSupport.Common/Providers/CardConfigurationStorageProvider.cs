// <copyright file="CardConfigurationStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Card configuration helps in fetching and storing dynamic card configuration in Azure table storage.
    /// </summary>
    public class CardConfigurationStorageProvider : StorageBaseProvider, ICardConfigurationStorageProvider
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CardConfigurationStorageProvider"/> class.
        /// </summary>
        /// <param name="configuration">The environment provided configuration.</param>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        public CardConfigurationStorageProvider(IConfiguration configuration, IOptionsMonitor<StorageOptions> storageOptions)
                        : base(storageOptions, Constants.CardConfigurationTable)
        {
            _ = this.InitializeDefaultCardTemplateAsync(configuration);
        }

        /// <summary>
        /// This method returns the latest Card template present in the Azure table storage.
        /// </summary>
        /// <returns>configuration details.</returns>
        public async Task<CardConfigurationEntity> GetConfigurationAsync()
        {
            await this.EnsureInitializedAsync();
            string filter = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, Constants.CardConfigurationPartitionKey);
            var query = new TableQuery<CardConfigurationEntity>().Where(filter);
            TableContinuationToken continuationToken = null;
            var configurations = new List<CardConfigurationEntity>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                configurations.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);

            return configurations.OrderByDescending(configuration => configuration.CreatedOn).FirstOrDefault();
        }

        /// <summary>
        /// Returns the latest card template present in the Azure table storage by CardId
        /// </summary>
        /// <param name="cardId">Unique identifier of the card configuration.</param>
        /// <returns>A <see cref="Task{TResult}"/>configuration details.</returns>
        public async Task<CardConfigurationEntity> GetConfigurationsByCardIdAsync(string cardId)
        {
            await this.EnsureInitializedAsync();
            string partitionFilterCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, Constants.CardConfigurationPartitionKey);
            string rowFilterCondition = TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, cardId);
            string filter = TableQuery.CombineFilters(partitionFilterCondition, TableOperators.And, rowFilterCondition);
            var query = new TableQuery<CardConfigurationEntity>().Where(filter);
            TableContinuationToken continuationToken = null;
            var configurations = new List<CardConfigurationEntity>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                configurations.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);

            return configurations.OrderByDescending(configuration => configuration.CreatedOn).FirstOrDefault();
        }

        /// <summary>
        /// Store or update card configuration entity in table storage.
        /// </summary>
        /// <param name="configurationEntity">Represents configuration entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents configuration entity is saved or updated.</returns>
        public async Task<CardConfigurationEntity> StoreOrUpdateEntityAsync(CardConfigurationEntity configurationEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(configurationEntity);
            var result = await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.Result as CardConfigurationEntity;
        }

        /// <summary>
        /// Returns the Adaptive card item element {Id, display name} mapping present in the Azure table storage by CardId
        /// </summary>
        /// <param name="cardId">Unique identifier of the card configuration.</param>
        /// <returns>A <see cref="Task{TResult}"/>configuration details.</returns>
        public async Task<Dictionary<string, string>> GetCardItemElementMappingAsync(string cardId)
        {
            Dictionary<string, string> cardElementMapping = new Dictionary<string, string>();
            CardConfigurationEntity configuration = await this.GetConfigurationsByCardIdAsync(cardId);
            var cardTemplates = JsonConvert.DeserializeObject<List<JObject>>(configuration?.CardTemplate);
            foreach (var template in cardTemplates)
            {
                var templateMapping = template.ToObject<AdaptiveCardPlaceHolderMapper>();
                cardElementMapping.Add(templateMapping.Id, templateMapping.DisplayName);
            }

            return cardElementMapping;
        }

        /// <summary>
        /// Initializes default adaptive card Json template to create new ticket request.
        /// </summary>
        /// <param name="configuration">Application configuration properties.</param>
        private async Task InitializeDefaultCardTemplateAsync(IConfiguration configuration)
        {
            await this.EnsureInitializedAsync();
            var card = await this.GetConfigurationAsync();
            if (card == null)
            {
                await this.StoreOrUpdateEntityAsync(new CardConfigurationEntity()
                {
                    CardId = Guid.NewGuid().ToString(),
                    CreatedOn = DateTime.UtcNow,
                    TeamId = Utility.ParseTeamIdFromDeepLink(configuration.GetValue<string>("Bot:TeamLink").ToString(CultureInfo.InvariantCulture)),
                    CardTemplate = configuration.GetValue<string>("Card:DefaultCardTemplate").ToString(CultureInfo.InvariantCulture),
                    TeamLink = configuration.GetValue<string>("Bot:TeamLink").ToString(CultureInfo.InvariantCulture),
                });
            }
        }
    }
}
