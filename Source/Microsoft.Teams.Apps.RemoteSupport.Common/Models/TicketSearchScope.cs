// <copyright file="TicketSearchScope.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    /// <summary>
    /// Class represents scope on basis of which tickets will be searched in messaging extension.
    /// </summary>
    public enum TicketSearchScope
    {
        /// <summary>
        /// Tickets with high priority.
        /// </summary>
        UrgentTickets,

        /// <summary>
        /// Tickets assigned to a subject-matter expert.
        /// </summary>
        AssignedTickets,

        /// <summary>
        /// Tickets which are not assigned to subject-matter expert.
        /// </summary>
        UnassignedTickets,

        /// <summary>
        /// Tickets which are active to subject-matter expert.
        /// </summary>
        ActiveTickets,

        /// <summary>
        /// Tickets which are not closed to subject-matter expert.
        /// </summary>
        ClosedTickets,
    }
}
