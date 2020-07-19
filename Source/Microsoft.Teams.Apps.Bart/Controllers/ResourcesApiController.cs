﻿// <copyright file="ResourcesApiController.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.Bart.Helpers;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.Error;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.Teams.Apps.Bart.Resources;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Meeting API controller for handling API calls made from react js client app (used in task module).
    /// </summary>
    [ApiController]
    [Route("api/[controller]/[action]")]
    //[Authorize]
    public class ResourcesApiController : ControllerBase
    {
        /// <summary>
        /// Telemetry client to log event and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Unauthorized error message response in case of user sign in failure.
        /// </summary>
        private const string SignInErrorCode = "signinRequired";

        /// <summary>
        /// Helper class to retrieve statuses.
        /// </summary>
        private readonly IStatusStorageProvider statusStorageProvider;

        /// <summary>
        /// Helper class to retrieve conference rooms.
        /// </summary>
        private readonly IConferenceBridgesStorageProvider conferenceBridgesStorageProvider;

        /// <summary>
        /// Generating and validating JWT token.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        private readonly IGraphApiHelper graphApiHelper;

        private readonly string graphApiToSearchUsers = "/v1.0/users?$filter=startswith(displayName,'{0}')&$select=displayName,userPrincipalName,id";

        private readonly string graphApiToGetIncidemntManagers = "/v1.0/groups/{0}/members?$select=displayName,userPrincipalName,id";

        /// <summary>
        /// Initializes a new instance of the <see cref="ResourcesApiController"/> class.
        /// </summary>
        /// <param name="statusStorageProvider">Helper class for getting available status.</param>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        /// <param name="conferenceBridgesStorageProvider">Helper class for getting available conference rooms.</param>
        public ResourcesApiController(ITokenHelper tokenHelper, IStatusStorageProvider statusStorageProvider, IConferenceBridgesStorageProvider conferenceBridgesStorageProvider, IGraphApiHelper graphApiHelper, TelemetryClient telemetryClient)
        {
            this.tokenHelper = tokenHelper;
            this.statusStorageProvider = statusStorageProvider;
            this.conferenceBridgesStorageProvider = conferenceBridgesStorageProvider;
            this.graphApiHelper = graphApiHelper;
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Get resource strings for displaying in client app.
        /// </summary>
        /// <returns>Object containing required strings.</returns>
        public ActionResult GetResourceStrings()
        {
            try
            {
                var strings = new
                {
                    Strings.Timezone,
                    Strings.SelectTimezone,
                    Strings.LoadingMessage,
                    Strings.MeetingLength,
                    Strings.SearchRoom,
                    Strings.BookRoom,
                    Strings.SearchRoomDropdownPlaceholder,
                    Strings.ExceptionResponse,
                    Strings.TimezoneNotSupported,
                    Strings.RoomUnavailable,
                    Strings.SelectDurationRoom,
                    Strings.Location,
                    Strings.AddButton,
                    Strings.DoneButton,
                    Strings.NoFavoriteRoomsTaskModule,
                    Strings.CantAddMoreRooms,
                    Strings.FavoriteRoomExist,
                    Strings.SelectRoomToAdd,
                    Strings.NoFavoritesDescriptionTaskModule,
                    Strings.SignInErrorMessage,
                    Strings.InvalidTenant,
                };
                return this.Ok(strings);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Get status and conference bridges for displaying in client app.
        /// </summary>
        /// <returns>Object containing required strings.</returns>
        public ActionResult GetAvailabilityData()
        {
            try
            {
                var bridges = this.conferenceBridgesStorageProvider.GetAsync().GetAwaiter().GetResult();
                return this.Ok(bridges.FindAll(bridge => bridge.Available));
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Get user data from AAD.
        /// </summary>
        /// <returns>Object containing list of users.</returns>
        public async Task<ActionResult> GetUsersAsync([FromQuery]int fromFlag, [FromQuery]string searchQuery)
        {
            try
            {
                var claims = this.GetUserClaims();
                this.telemetryClient.TrackTrace($"User {claims.UserObjectIdentifer} submitted request to get supported time zones.");

                var token = await this.tokenHelper.GetUserTokenAsync("29:1gMQTXLxN-dQImkrSeGvgGSvxl4VTOaSuwdqnH8RuWvysIlFR3rJRwy6vZGmiiR3BDHzJUZxDpegnBNWhbNGFTw").ConfigureAwait(false);
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

                string url = string.Format(this.graphApiToGetIncidemntManagers, "b423a73a-e033-4c91-9bd8-8f45a14a56da");
                if (fromFlag == 1)
                {
                    url = string.Format(this.graphApiToSearchUsers, searchQuery);
                }

                var result = await this.graphApiHelper.GetAsync(url, token).ConfigureAwait(false);
                var responseMessage = await result.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!string.IsNullOrEmpty(responseMessage))
                {
                    var list = JObject.Parse(responseMessage)["value"];
                    return this.Ok(JsonConvert.DeserializeObject<List<User>>(list.ToString()));
                }

                this.telemetryClient.TrackTrace($"No results found for the search query.");
                return this.StatusCode(
                    StatusCodes.Status404NotFound,
                    new Error
                    {
                        StatusCode = "Not found",
                        ErrorMessage = "No results found for the search query.",
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