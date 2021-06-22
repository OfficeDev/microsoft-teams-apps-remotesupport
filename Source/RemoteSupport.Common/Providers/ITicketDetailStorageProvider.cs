// <copyright file="ITicketDetailStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;

    /// <summary>
    /// Ticket provider helps in fetching and storing information in storage table.
    /// </summary>
    public interface ITicketDetailStorageProvider
    {
        /// <summary>
        /// Save or update ticket entity.
        /// </summary>
        /// <param name="ticketDetails">Ticket received from bot based on which appropriate row will replaced or inserted in table storage.</param>
        /// <returns><see cref="Task"/> that resolves successfully if the data was saved successfully.</returns>
        Task<bool> UpsertTicketAsync(TicketDetail ticketDetails);

        /// <summary>
        /// Get already saved entity detail from storage table.
        /// </summary>
        /// <param name="ticketId">ticket id received from bot based on which appropriate row data will be fetched.</param>
        /// <returns><see cref="Task"/> Already saved entity detail.</returns>
        Task<TicketDetail> GetTicketAsync(string ticketId);
    }
}
