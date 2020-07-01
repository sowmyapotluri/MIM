// <copyright file="IncidentApiController.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.Bart.Helpers;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.Error;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.Teams.Apps.Bart.Providers.Storage;
    using TimeZoneConverter;

    /// <summary>
    /// Meeting API controller for handling API calls made from react js client app (used in task module).
    /// </summary>
    [ApiController]
    [Route("api/[controller]/[action]")]
    //[Authorize]
    public class IncidentApiController : ControllerBase
    {
        /// <summary>
        /// Number of rooms to load in dropdown initially.
        /// </summary>
        private const int InitialRoomCount = 5;

        /// <summary>
        /// Unauthorized error message response in case of user sign in failure.
        /// </summary>
        private const string SignInErrorCode = "signinRequired";

        /// <summary>
        /// Telemetry client to log event and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Generating and validating JWT token.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Helper class which exposes methods required for incident creation.
        /// </summary>
        private readonly IServiceNowProvider serviceNowProvider;

        /// <summary>
        /// Helper class which exposes methods required for incident creation.
        /// </summary>
        private readonly IConferenceBridgesStorageProvider conferenceBridgesStorageProvider;

        /// <summary>
        /// Storage provider to perform fetch operation on UserConfiguration table.
        /// </summary>
        private readonly IUserConfigurationStorageProvider userConfigurationStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="IncidentApiController"/> class.
        /// </summary>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        /// <param name="tokenHelper">Generating and validating JWT token.</param>
        /// <param name="serviceNowProvider">Helper class which exposes methods required for incident creation.</param>
        /// <param name="conferenceBridgesStorageProvider">Helper class which exposes methods required for getting and updating conference room status.</param>
        /// <param name="userConfigurationStorageProvider">Storage provider to perform fetch operation on UserConfiguration table.</param>
        public IncidentApiController(TelemetryClient telemetryClient, ITokenHelper tokenHelper, IServiceNowProvider serviceNowProvider, IConferenceBridgesStorageProvider conferenceBridgesStorageProvider, IUserConfigurationStorageProvider userConfigurationStorageProvider)
        {
            this.telemetryClient = telemetryClient;
            this.tokenHelper = tokenHelper;
            this.serviceNowProvider = serviceNowProvider;
            this.conferenceBridgesStorageProvider = conferenceBridgesStorageProvider;
            this.userConfigurationStorageProvider = userConfigurationStorageProvider;
        }

        /// <summary>
        /// Get supported time zones for user from Graph API.
        /// </summary>
        /// <param name="incident">Incident object.</param>
        /// <returns>Returns the newly created incident data.</returns>
        [HttpPost]
        public async Task<IActionResult> CreateIncidentAsync([FromBody]Incident incident)
        {
            try
            {
                var claims = this.GetUserClaims();
                this.telemetryClient.TrackTrace($"User {claims.UserObjectIdentifer} submitted request to get supported time zones.");

                var token = await this.tokenHelper.GetUserTokenAsync(claims.FromId).ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    this.telemetryClient.TrackTrace($"Azure Active Directory access token for user {claims.UserObjectIdentifer} is empty.");
                    return this.StatusCode(
                        StatusCodes.Status401Unauthorized,
                        new Error
                        {
                            StatusCode = SignInErrorCode,
                            ErrorMessage = "Azure Active Directory access token for user is found empty.",
                        });
                }

                var bridgeStatus = await this.conferenceBridgesStorageProvider.GetAsync(incident.Bridge.Code).ConfigureAwait(false);
                if (bridgeStatus.Available)
                {
                    Incident incidentCreated = await this.serviceNowProvider.CreateIncidentAsync(incident, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa");
                    //if (string.IsNullOrEmpty(incident.Id))
                    //{
                    //    incident.Bridge.Available = false;
                    //    await this.conferenceBridgesStorageProvider.AddAsync(incident.Bridge).ConfigureAwait(false);
                    //}

                    return this.Ok(incidentCreated);
                }

                return this.StatusCode(
                    StatusCodes.Status409Conflict,
                    new Error
                    {
                        StatusCode = "Confilt",
                        ErrorMessage = "Bridge not available.",
                    });
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }


        /// <summary>
        /// Get claims of user.
        /// </summary>
        /// <returns>Claims.</returns>
        private JwtClaim GetUserClaims()
        {
            var claims = this.User.Claims;
            var jwtClaims = new JwtClaim
            {
                FromId = claims.Where(claim => claim.Type == "fromId").Select(claim => claim.Value).First(),
                ServiceUrl = claims.Where(claim => claim.Type == "serviceURL").Select(claim => claim.Value).First(),
                UserObjectIdentifer = claims.Where(claim => claim.Type == "userObjectIdentifer").Select(claim => claim.Value).First(),
            };

            return jwtClaims;
        }
    }
}