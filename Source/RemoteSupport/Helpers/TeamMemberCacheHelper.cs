// <copyright file="TeamMemberCacheHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Helpers
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Caching.Memory;

    /// <summary>
    /// Class that handles the card configuration.
    /// </summary>
    public static class TeamMemberCacheHelper
    {
        /// <summary>
        /// Cache key for expert details
        /// </summary>
        private const string ExpertCollectionCacheKey = "_expertCollectionKey";

        /// <summary>
        /// Sets the team members cache duration.
        /// </summary>
        private static readonly TimeSpan CacheDuration = TimeSpan.FromDays(1);

        /// <summary>
        /// Provide team members information.
        /// </summary>
        /// <param name="memoryCache">MemoryCache instance for caching on call expert objectId's.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="userId">Describes a user Id.</param>
        /// <param name="teamId">Describes a team Id.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Returns team members information from cache.</returns>
        public static async Task<TeamsChannelAccount> GetMemberInfoAsync(IMemoryCache memoryCache, ITurnContext turnContext, string userId, string teamId, CancellationToken cancellationToken)
        {
            bool isCacheEntryExists = memoryCache.TryGetValue(ExpertCollectionCacheKey + userId, out TeamsChannelAccount memberInformation);

            if (!isCacheEntryExists)
            {
                if (teamId != null)
                {
                    memberInformation = await TeamsInfo.GetTeamMemberAsync(turnContext, userId, teamId);
                }
                else
                {
                    memberInformation = await TeamsInfo.GetMemberAsync(turnContext, userId, cancellationToken);
                }

                if (memberInformation != null)
                {
                    memoryCache.Set(ExpertCollectionCacheKey + userId, memberInformation, CacheDuration);
                }
            }

            return memberInformation;
        }
    }
}
