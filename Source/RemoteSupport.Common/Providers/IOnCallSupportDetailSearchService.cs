// <copyright file="IOnCallSupportDetailSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;

    /// <summary>
    /// Interface to provide Search on call support team based on search query.
    /// </summary>
    public interface IOnCallSupportDetailSearchService
    {
        /// <summary>
        /// Provide search result for table to be used by SME based on Azure search service.
        /// </summary>
        /// <param name="searchQuery">searchQuery to be provided by message extension.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <returns>List of search results.</returns>
        Task<IEnumerable<OnCallSupportDetail>> SearchOnCallSupportTeamAsync(string searchQuery, int? count = null, int? skip = null);
    }
}
