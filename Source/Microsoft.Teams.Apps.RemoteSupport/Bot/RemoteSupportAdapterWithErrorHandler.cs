// <copyright file="RemoteSupportAdapterWithErrorHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Bot
{
    using System;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Implements Error Handler.
    /// </summary>
    public class RemoteSupportAdapterWithErrorHandler : BotFrameworkHttpAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RemoteSupportAdapterWithErrorHandler"/> class.
        /// </summary>
        /// <param name="configuration">Application configurations.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="remoteSupportActivityMiddleWare">Represents middle ware that can operate on incoming activities.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="conversationState">conversationState.</param>
        public RemoteSupportAdapterWithErrorHandler(IConfiguration configuration, ILogger<IBotFrameworkHttpAdapter> logger, RemotesupportActivityMiddleWare remoteSupportActivityMiddleWare, IStringLocalizer<Strings> localizer, ConversationState conversationState = null)
            : base(configuration)
        {
            if (remoteSupportActivityMiddleWare == null)
            {
                throw new NullReferenceException(nameof(remoteSupportActivityMiddleWare));
            }

            // Add activity middle ware to the adapter's middle ware pipeline
            this.Use(remoteSupportActivityMiddleWare);

            this.OnTurnError = async (turnContext, exception) =>
            {
                // Log any leaked exception from the application.
                logger.LogError(exception, $"Exception caught : {exception.Message}");

                // Send a catch-all apology to the user.
                await turnContext.SendActivityAsync(localizer.GetString("ErrorMessage"));

                if (conversationState != null)
                {
                    try
                    {
                        // Delete the conversationState for the current conversation to prevent the
                        // bot from getting stuck in a error-loop caused by being in a bad state.
                        // ConversationState should be thought of as similar to "cookie-state" in a Web pages.
                        await conversationState.DeleteAsync(turnContext);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, $"Exception caught on attempting to Delete ConversationState : {ex.Message}");
                    }
                }
            };
        }
    }
}
