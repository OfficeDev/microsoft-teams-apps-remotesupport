// <copyright file="ITicketIdGeneratorStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Providers
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;

    /// <summary>
    /// Interface to generate ticket ids from table storage.
    /// </summary>
    public interface ITicketIdGeneratorStorageProvider
    {
        /// <summary>
        /// Get a new ticket id from the table storage.
        /// </summary>
        /// <returns>Next TicketId generated from table storage.</returns>
        Task<int> GetTicketIdAsync();

        /// <summary>
        /// update the ticket id in the table storage.
        /// </summary>
        /// <param name="ticketIdGenerator"> Entity containing latest ticket Id details.</param>
        /// <returns> Returns next ticket id generated from table storage.</returns>
        Task<int> UpdateTicketIdAsync(TicketIdGenerator ticketIdGenerator);
    }
}
