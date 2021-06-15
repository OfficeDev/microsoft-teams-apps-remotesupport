// <copyright file="AzureAdClaimTypes.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common.Models
{
    /// <summary>
    /// Azure Active Directory Claim Types.
    /// </summary>
    public static class AzureAdClaimTypes
    {
        /// <summary>
        /// Object Identifier.
        /// </summary>
        public const string ObjectId = "http://schemas.microsoft.com/identity/claims/objectidentifier";

        /// <summary>
        /// Identity Scope.
        /// </summary>
        public const string Scope = "http://schemas.microsoft.com/identity/claims/scope";
    }
}