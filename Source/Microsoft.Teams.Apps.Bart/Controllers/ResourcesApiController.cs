// <copyright file="ResourcesApiController.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Controllers
{
    using System;
    using System.Collections.Generic;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.Teams.Apps.Bart.Resources;

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
        /// Helper class to retrieve statuses.
        /// </summary>
        private readonly IStatusStorageProvider statusStorageProvider;

        /// <summary>
        /// Helper class to retrieve conference rooms.
        /// </summary>
        private readonly IConferenceBridgesStorageProvider conferenceBridgesStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResourcesApiController"/> class.
        /// </summary>
        /// <param name="statusStorageProvider">Helper class for getting available status.</param>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        /// <param name="conferenceBridgesStorageProvider">Helper class for getting available conference rooms.</param>
        public ResourcesApiController(IStatusStorageProvider statusStorageProvider, IConferenceBridgesStorageProvider conferenceBridgesStorageProvider, TelemetryClient telemetryClient)
        {
            this.statusStorageProvider = statusStorageProvider;
            this.conferenceBridgesStorageProvider = conferenceBridgesStorageProvider;
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
    }
}