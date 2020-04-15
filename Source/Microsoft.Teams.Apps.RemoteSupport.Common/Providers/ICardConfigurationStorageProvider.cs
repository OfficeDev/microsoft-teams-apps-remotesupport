// <copyright file="ICardConfigurationStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;

    /// <summary>
    /// Card configuration helps in fetching and storing dynamic card configuration in Azure table storage.
    /// </summary>
    public interface ICardConfigurationStorageProvider
    {
        /// <summary>
        /// Stores configuration details into Azure table storage.
        /// </summary>
        /// <param name="configurationEntity">Configuration storage entity.</param>
        /// <returns>A task that represents configuration entity is saved or updated.</returns>
        Task<CardConfigurationEntity> StoreOrUpdateEntityAsync(CardConfigurationEntity configurationEntity);

        /// <summary>
        /// Returns the latest card template present in the Azure table storage.
        /// </summary>
        /// <returns>configuration details.</returns>
        Task<CardConfigurationEntity> GetConfigurationAsync();

        /// <summary>
        /// Returns the latest card template present in the Azure table storage by CardId
        /// </summary>
        /// <param name="cardId">Unique identifier of the card configuration.</param>
        /// <returns>A <see cref="Task{TResult}"/>configuration details.</returns>
        Task<CardConfigurationEntity> GetConfigurationsByCardIdAsync(string cardId);

        /// <summary>
        /// Returns the Adaptive card item element {Id, display name} mapping present in the Azure table storage by CardId
        /// </summary>
        /// <param name="cardId">Unique identifier of the card configuration.</param>
        /// <returns>A <see cref="Task{TResult}"/>configuration details.</returns>
        Task<Dictionary<string, string>> GetCardItemElementMappingAsync(string cardId);
    }
}
