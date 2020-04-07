// <copyright file="AzureAdSettings.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Configuration.Models
{
    /// <summary>
    /// Azure AD configuration model.
    /// </summary>
    public class AzureAdSettings
    {
        /// <summary>
        /// Gets or sets Client Id.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Gets or sets the Azure AD instance.
        /// </summary>
        public string Instance { get; set; }

        /// <summary>
        /// Gets or sets Tenant Id.
        /// </summary>
        public string Tenant { get; set; }

        /// <summary>
        /// Gets or sets User Principal Name.
        /// </summary>
        public string Upn { get; set; }
    }
}