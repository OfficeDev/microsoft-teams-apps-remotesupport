// <copyright file="RemoteSupportActivityHandlerOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport
{
    /// <summary>
    /// The RemoteSupportActivityHandlerOptions are the options for the <see cref="RemoteSupportActivityHandler" /> bot.
    /// </summary>
    public sealed class RemoteSupportActivityHandlerOptions
    {
        /// <summary>
        /// Gets or sets a value indicating whether the response to a message should be all uppercase.
        /// </summary>
        public bool UpperCaseResponse { get; set; }

        /// <summary>
        /// Gets or sets unique id of Tenant.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets unique identifier of team in which Bot is installed.
        /// </summary>
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets application base Uri.
        /// </summary>
        public string AppBaseUri { get; set; }
    }
}