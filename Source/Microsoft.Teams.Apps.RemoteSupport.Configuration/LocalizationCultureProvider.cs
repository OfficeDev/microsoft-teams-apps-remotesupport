// <copyright file="LocalizationCultureProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Localization;

    /// <summary>
    /// The BotLocalizationCultureProvider is responsible for implementing the <see cref="IRequestCultureProvider"/> for Bot Activities
    /// received from BotFramework.
    /// </summary>
    internal sealed class LocalizationCultureProvider : IRequestCultureProvider
    {
        /// <summary>
        /// Get the culture of the current request.
        /// </summary>
        /// <param name="httpContext">The current request.</param>
        /// <returns>A Task resolving to the culture info if found, null otherwise.</returns>
        #pragma warning disable UseAsyncSuffix // Interface method doesn't have Async suffix.
        public async Task<ProviderCultureResult> DetermineProviderCultureResult(HttpContext httpContext)
        #pragma warning restore UseAsyncSuffix
        {
            if (httpContext?.Request?.Body?.CanRead != true)
            {
                return null;
            }

            try
            {
                // Wrap the request stream so that we can rewind it back to the start for regular request processing.
                httpContext.Request.EnableBuffering();

                // Read the request headers and set the parsed culture information.
                var locale = httpContext.Request.Headers["Accept-Language"].First()?.Split(',').First();
                var result = new ProviderCultureResult(locale);
                return result;
            }
            #pragma warning disable CA1031 // part of the middle ware pipeline, better to use default local then fail the request.
            catch (Exception)
            #pragma warning restore CA1031
            {
                return null;
            }
            finally
            {
                httpContext.Request.Body.Seek(0, SeekOrigin.Begin);
            }
        }
    }
}