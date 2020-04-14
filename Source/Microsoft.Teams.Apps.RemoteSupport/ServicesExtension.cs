// <copyright file="ServicesExtension.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using Microsoft.AspNetCore.Authentication.JwtBearer;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Localization;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Providers;
    using Microsoft.Teams.Apps.RemoteSupport.Helpers;

    /// <summary>
    /// Class to extend ServiceCollection .
    /// </summary>
    public static class ServicesExtension
    {
        /// <summary>
        /// Adds application configuration settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddConfigurationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            services.Configure<RemoteSupportActivityHandlerOptions>(options =>
            {
                options.UpperCaseResponse = configuration.GetValue<bool>("UppercaseResponse");
                options.TenantId = configuration.GetValue<string>("Bot:TenantId");

                // Parse team Id from deep link.
                options.TeamId = Common.Utility.ParseTeamIdFromDeepLink(configuration.GetValue<string>("Bot:TeamLink"));
                options.AppBaseUri = configuration.GetValue<string>("Bot:AppBaseUri");
            });
            services.Configure<TokenOptions>(options =>
            {
                options.SecurityKey = configuration.GetValue<string>("Token:SecurityKey");
            });
            services.Configure<StorageOptions>(options =>
            {
                options.ConnectionString = configuration.GetValue<string>("Storage:ConnectionString");
            });
            services.Configure<SearchServiceOptions>(options =>
            {
                options.SearchServiceName = configuration.GetValue<string>("Search:SearchServiceName");
                options.SearchServiceQueryApiKey = configuration.GetValue<string>("Search:SearchServiceQueryApiKey");
                options.SearchServiceAdminApiKey = configuration.GetValue<string>("Search:SearchServiceAdminApiKey");
                options.SearchIndexingIntervalInMinutes = configuration.GetValue<string>("Search:SearchIndexingIntervalInMinutes");
            });
        }

        /// <summary>
        /// Adds providers to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        public static void AddProviders(this IServiceCollection services)
        {
            services.AddSingleton<ITicketDetailStorageProvider, TicketDetailStorageProvider>();
            services.AddSingleton<IOnCallSupportDetailStorageProvider, OnCallSupportDetailStorageProvider>();
            services.AddSingleton<ITicketSearchService, TicketSearchService>();
            services.AddSingleton<ITicketIdGeneratorStorageProvider, TicketIdGeneratorStorageProvider>();
            services.AddSingleton<ICardConfigurationStorageProvider, CardConfigurationStorageProvider>();
            services.AddSingleton<IOnCallSupportDetailSearchService, OnCallSupportDetailSearchService>();
            services.AddSingleton<ITokenHelper, TokenHelper>();
        }

        /// <summary>
        /// Adds custom JWT authentication to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddCustomJWTAuthentication(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
                .AddJwtBearer(options =>
                {
                    options.TokenValidationParameters = new TokenValidationParameters
                    {
                        ValidateAudience = true,
                        ValidAudiences = new List<string> { configuration.GetValue<string>("Bot:AppBaseUri") },
                        ValidIssuers = new List<string> { configuration.GetValue<string>("Bot:AppBaseUri") },
                        ValidateIssuer = true,
                        ValidateIssuerSigningKey = true,
                        IssuerSigningKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(configuration.GetValue<string>("Token:SecurityKey"))),
                        RequireExpirationTime = true,
                        ValidateLifetime = true,
                        ClockSkew = TimeSpan.FromSeconds(30),
                    };
                });
        }

        /// <summary>
        /// Adds credential providers for authentication.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddCredentialProviders(this IServiceCollection services, IConfiguration configuration)
        {
            services
                .AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
            services.AddSingleton(new MicrosoftAppCredentials(configuration.GetValue<string>("MicrosoftAppId"), configuration.GetValue<string>("MicrosoftAppPassword")));
        }

        /// <summary>
        /// Adds localization settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddLocalizationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddLocalization(options => options.ResourcesPath = "Resources");
            services.Configure<RequestLocalizationOptions>(options =>
            {
                var defaultCulture = CultureInfo.GetCultureInfo(configuration.GetValue<string>("i18n:DefaultCulture"));
                var supportedCultures = configuration.GetValue<string>("i18n:SupportedCultures").Split(',')
                    .Select(culture => CultureInfo.GetCultureInfo(culture))
                    .ToList();

                options.DefaultRequestCulture = new RequestCulture(defaultCulture);
                options.SupportedCultures = supportedCultures;
                options.SupportedUICultures = supportedCultures;

                options.RequestCultureProviders = new List<IRequestCultureProvider>
                {
                    new BotLocalizationCultureProvider(),
                };
            });
        }
    }
}
