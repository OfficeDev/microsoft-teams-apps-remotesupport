// <copyright file="TokenHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.IdentityModel.Tokens.Jwt;
    using System.Security.Claims;
    using System.Text;
    using Microsoft.Extensions.Options;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.RemoteSupport;
    using Microsoft.Teams.Apps.RemoteSupport.Common.Models;

    /// <summary>
    /// Helper class for JWT token generation and validation.
    /// </summary>
    public class TokenHelper : ITokenHelper
    {
        /// <summary>
        /// Security key for generating and validating token.
        /// </summary>
        private readonly string securityKey;

        /// <summary>
        /// Application base URL.
        /// </summary>
        private readonly string appBaseUri;

        /// <summary>
        /// Initializes a new instance of the <see cref="TokenHelper"/> class.
        /// </summary>
        /// <param name="remoteSupportActivityHandlerOptions">A set of key/value application configuration properties for Remote Support bot.</param>
        /// <param name="tokenOptions">A set of key/value application configuration properties for token.</param>
        public TokenHelper(
            IOptionsMonitor<RemoteSupportActivityHandlerOptions> remoteSupportActivityHandlerOptions,
            IOptionsMonitor<TokenOptions> tokenOptions)
        {
            tokenOptions = tokenOptions ?? throw new ArgumentNullException(nameof(tokenOptions));
            remoteSupportActivityHandlerOptions = remoteSupportActivityHandlerOptions ?? throw new ArgumentNullException(nameof(remoteSupportActivityHandlerOptions));
            this.securityKey = tokenOptions.CurrentValue.SecurityKey;
            this.appBaseUri = remoteSupportActivityHandlerOptions.CurrentValue.AppBaseUri;
        }

        /// <summary>
        /// Generate JWT token used by client application to authenticate HTTP calls with API.
        /// </summary>
        /// <param name="applicationBasePath">Service URL from bot.</param>
        /// <param name="fromId">Unique Id from activity.</param>
        /// <param name="jwtExpiryMinutes">Expiry of token.</param>
        /// <returns>JWT token.</returns>
        public string GenerateAPIAuthToken(string applicationBasePath, string fromId, int jwtExpiryMinutes)
        {
            SymmetricSecurityKey signingKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(this.securityKey));
            SigningCredentials signingCredentials = new SigningCredentials(signingKey, SecurityAlgorithms.HmacSha256);

            SecurityTokenDescriptor securityTokenDescriptor = new SecurityTokenDescriptor()
            {
                Subject = new ClaimsIdentity(
                    new List<Claim>()
                    {
                        new Claim("applicationBasePath", applicationBasePath),
                        new Claim("fromId", fromId),
                    }, "Custom"),
                NotBefore = DateTime.UtcNow,
                SigningCredentials = signingCredentials,
                Issuer = this.appBaseUri,
                Audience = this.appBaseUri,
                IssuedAt = DateTime.UtcNow,
                Expires = DateTime.UtcNow.AddMinutes(jwtExpiryMinutes),
            };

            JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
            SecurityToken token = tokenHandler.CreateToken(securityTokenDescriptor);
            return tokenHandler.WriteToken(token);
        }
    }
}
