// <copyright file="Utility.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Common
{
    using System;
    using System.Text.RegularExpressions;
    using System.Web;

    /// <summary>
    /// Utility class for common functionality.
    /// </summary>
    public static class Utility
    {
        /// <summary>
        /// Based on deep link URL received find team id and set it.
        /// </summary>
        /// <param name="teamIdDeepLink">Deep link to get the team id.</param>
        /// <returns>A team id from the deep link URL.</returns>
        public static string ParseTeamIdFromDeepLink(string teamIdDeepLink)
        {
            // team id regex match
            // for a pattern like https://teams.microsoft.com/l/team/19%3a64c719819fb1412db8a28fd4a30b581a%40thread.tacv2/conversations?groupId=53b4782c-7c98-4449-993a-441870d10af9&tenantId=72f988bf-86f1-41af-91ab-2d7cd011db47
            // regex checks for 19%3a64c719819fb1412db8a28fd4a30b581a%40thread.tacv2
            var match = Regex.Match(teamIdDeepLink, @"teams.microsoft.com/l/team/(\S+)/");
            if (!match.Success)
            {
                throw new ArgumentException($"Invalid team found.");
            }

            return HttpUtility.UrlDecode(match.Groups[1].Value);
        }
    }
}