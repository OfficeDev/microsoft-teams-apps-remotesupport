// <copyright file="BaseRemoteSupportController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RemoteSupport.Controllers
{
    using System.Linq;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.RemoteSupport.Models;

    /// <summary>
    /// Base controller to handle user and company response API operations.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class BaseRemoteSupportController : ControllerBase
    {
        /// <summary>
        /// Get claims of user.
        /// </summary>
        /// <returns>User claims.</returns>
        protected JwtClaims GetUserClaims()
        {
            var claims = this.User.Claims;
            var jwtClaims = new JwtClaims
            {
                FromId = claims.Where(claim => claim.Type == "fromId").Select(claim => claim.Value).First(),
                ApplicationBasePath = claims.Where(claim => claim.Type == "applicationBasePath").Select(claim => claim.Value).First(),
            };

            return jwtClaims;
        }

        /// <summary>
        /// Method checks if user is authorized to make API calls.
        /// </summary>
        /// <returns> Returns success/failure depending on whether user is authorized.</returns>
        protected bool IsUserAuthenticated()
        {
            var fromId = this.User.Claims.Where(claim => claim.Type == "fromId").Select(claim => claim.Value).FirstOrDefault();
            if (string.IsNullOrEmpty(fromId))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Creates the error response as per the status codes in case of error.
        /// </summary>
        /// <param name="statusCode">Describes the type of error.</param>
        /// <param name="errorMessage">Describes the error message.</param>
        /// <returns>Returns error response with appropriate message and status code.</returns>
        protected IActionResult GetErrorResponse(int statusCode, string errorMessage)
        {
            switch (statusCode)
            {
                case StatusCodes.Status401Unauthorized:
                    return this.StatusCode(
                        StatusCodes.Status401Unauthorized,
                        new ErrorResponse
                        {
                            StatusCode = "signinRequired",
                            ErrorMessage = errorMessage,
                        });
                case StatusCodes.Status400BadRequest:
                    return this.StatusCode(
                        StatusCodes.Status400BadRequest,
                        new ErrorResponse
                        {
                            StatusCode = "badRequest",
                            ErrorMessage = errorMessage,
                        });
                default:
                    return this.StatusCode(
                        StatusCodes.Status500InternalServerError,
                        new ErrorResponse
                        {
                            StatusCode = "internalServerError",
                            ErrorMessage = errorMessage,
                        });
            }
        }
    }
}