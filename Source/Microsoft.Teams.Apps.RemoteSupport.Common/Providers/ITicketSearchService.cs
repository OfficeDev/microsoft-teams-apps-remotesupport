// <copyright file="ITicketSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;

    /// <summary>
    /// Interface of search service which will help in creating index, indexer and data source if it doesn't exist.
    /// </summary>
    public interface ITicketSearchService
    {
        /// <summary>
        /// Provide search result for table to be used by SME based on Azure search service.
        /// </summary>
        /// <param name="searchScope">Scope of the search.</param>
        /// <param name="searchQuery">searchQuery to be provided by message extension.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="requestorId">Requester id of the user to get specific tickets.</param>
        /// <returns>List of search results.</returns>
        Task<IList<TicketDetail>> SearchTicketsAsync(TicketSearchScope searchScope, string searchQuery, int? count = null, int? skip = null, string requestorId = "");
    }
}
