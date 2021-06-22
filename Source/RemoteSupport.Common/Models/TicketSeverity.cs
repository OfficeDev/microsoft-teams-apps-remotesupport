// <copyright file="TicketSeverity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    /// <summary>
    /// Represents the current severity of ticket.
    /// </summary>
    public enum TicketSeverity
    {
        /// <summary>
        /// Represents that ticket needs to be addressed on normal priority.
        /// </summary>
        Normal = 0,

        /// <summary>
        /// Represents that ticket needs to be addressed on high priority.
        /// </summary>
        Urgent = 1,
    }
}
