// <copyright file="IOnCallSupportDetailStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;

    /// <summary>
    /// Interface for on call support detail provider.
    /// </summary>
    public interface IOnCallSupportDetailStorageProvider
    {
        /// <summary>
        /// Save on call support details in Azure Table Storage.
        /// </summary>
        /// <param name="onCallSupportTeamDetails">On call support details to be stored in table storage.</param>
        /// <returns><see cref="Task"/> Returns OnCallSupportId when on call support data was saved successfully.</returns>
        Task<string> UpsertOnCallSupportDetailsAsync(OnCallSupportDetail onCallSupportTeamDetails);

        /// <summary>
        /// Get already saved entity detail from storage table.
        /// </summary>
        /// <param name="onCallSMEId">onCallSMEId received from bot based on which appropriate row data will be fetched.</param>
        /// <returns><see cref="Task"/> Already saved entity detail.</returns>
        Task<OnCallSupportDetail> GetOnCallSupportDetailAsync(string onCallSMEId);
    }
}
