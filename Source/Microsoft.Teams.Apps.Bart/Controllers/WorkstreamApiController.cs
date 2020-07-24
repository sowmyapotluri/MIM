// <copyright file="WorkstreamApiController.cs" company="Microsoft Corporation">
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
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.Bart.Cards;
    using Microsoft.Teams.Apps.Bart.Helpers;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.Error;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.Teams.Apps.Bart.Providers.Storage;
    using Newtonsoft.Json;
    using TimeZoneConverter;

    /// <summary>
    /// Meeting API controller for handling API calls made from react js client app (used in task module).
    /// </summary>
    [ApiController]
    [Route("api/[controller]/[action]")]
    //[Authorize]
    public class WorkstreamApiController : ControllerBase
    {

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
        /// Storage provider to perform creation and updation on Workstreans table.
        /// </summary>
        private readonly IWorkstreamStorageProvider workstreamStorageProvider;

        /// <summary>
        /// Storage provider to perform fetch operation on Incident table .
        /// </summary>
        private readonly IIncidentStorageProvider incidentStorageProvider;

        /// <summary>
        /// Helper class which exposes methods required for incident creation and updation.
        /// </summary>
        private readonly IServiceNowProvider serviceNowProvider;

        /// <summary>
        /// Storage provider to perform fetch operation on UserConfiguration table.
        /// </summary>
        private readonly IUserConfigurationStorageProvider userConfigurationStorageProvider;

        /// <summary>
        /// App credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Initializes a new instance of the <see cref="WorkstreamApiController"/> class.
        /// </summary>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        /// <param name="tokenHelper">Generating and validating JWT token.</param>
        /// <param name="userConfigurationStorageProvider">Storage provider to perform fetch operation on UserConfiguration table.</param>
        /// <param name="workstreamStorageProvider">Helper class which exposes methods required for workstream creation.</param>
        /// <param name="incidentStorageProvider">Storage provider to perform fetch operation on Incident table .</param>
        /// <param name="microsoftAppCredentials">App credentials.</param>
        /// <param name="serviceNowProvider">Helper class which exposes methods required for incident creation and updation.</param>

        public WorkstreamApiController(TelemetryClient telemetryClient, ITokenHelper tokenHelper, IUserConfigurationStorageProvider userConfigurationStorageProvider, 
            IWorkstreamStorageProvider workstreamStorageProvider, IIncidentStorageProvider incidentStorageProvider, MicrosoftAppCredentials microsoftAppCredentials,
            IServiceNowProvider serviceNowProvider)
        {
            this.telemetryClient = telemetryClient;
            this.tokenHelper = tokenHelper;
            this.userConfigurationStorageProvider = userConfigurationStorageProvider;
            this.workstreamStorageProvider = workstreamStorageProvider;
            this.incidentStorageProvider = incidentStorageProvider;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.serviceNowProvider = serviceNowProvider;
        }

        /// <summary>
        /// Create workstream associated with an incident.
        /// </summary>
        /// <param name="workstreams">List of workstream object.</param>
        /// <returns>Returns status after creation.</returns>
        [HttpPost]
        public async Task<IActionResult> CreateOrUpdateWorkstremAsync([FromBody]List<WorkstreamEntity> workstreams)
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
                        new Models.Error.Error
                        {
                            StatusCode = SignInErrorCode,
                            ErrorMessage = "Azure Active Directory access token for user is found empty.",
                        });
                }

                var sentWorkstreamNotificationTask = new List<Task>();
                if (workstreams.Count > 0)
                {
                    List<string> workstreamString = new List<string>();
                    var incidentTableData = await this.incidentStorageProvider.GetAsync(workstreams.FirstOrDefault().PartitionKey).ConfigureAwait(false);
                    Incident incident = await this.serviceNowProvider.GetIncidentAsync(incidentTableData.RowKey, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                    incident.BridgeDetails = new ConferenceRoomEntity
                    {
                        Code = incidentTableData.BridgeId,
                        BridgeURL = incidentTableData.BridgeLink,
                    };
                    foreach (var workstream in workstreams)
                    {
                        if (workstream.InActive)
                        {
                            workstream.ETag = "*";
                            await this.workstreamStorageProvider.DeleteAsync(workstream).ConfigureAwait(false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(workstream.Description))
                            {
                                if (string.IsNullOrEmpty(workstream.Id))
                                {
                                    workstreams.Remove(workstream);
                                }

                                if (workstream.New)
                                {
                                    var user = await this.userConfigurationStorageProvider.GetAsync(workstream.AssignedToId).ConfigureAwait(false);
                                    if (user != null)
                                    {
                                        MicrosoftAppCredentials.TrustServiceUrl(incidentTableData.ServiceUrl);
                                        var connector = new ConnectorClient(new Uri(incidentTableData.ServiceUrl), this.microsoftAppCredentials);
                                        var card = new IncidentCard(incident).GetIncidentAttachment();

                                        // Sending cards to team and personal chat
                                        sentWorkstreamNotificationTask.Add(connector.Conversations.SendToConversationAsync(user.ConversationId, (Activity)MessageFactory.Attachment(card)));
                                    }
                                }

                                workstream.New = false;
                                await this.workstreamStorageProvider.AddAsync(workstream).ConfigureAwait(false);
                            }
                        }
                    }

                    incident.WorkNotes = string.Join(',', workstreamString);
                    await this.serviceNowProvider.UpdateIncidentAsync(incident, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                    await Task.WhenAll(sentWorkstreamNotificationTask).ConfigureAwait(false);
                }

                this.telemetryClient.TrackEvent($"Workstreams entered into database - Incident Number: {workstreams.FirstOrDefault().PartitionKey} && workstream count= {workstreams.Count}");

                return this.Ok();
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Get all workstream associated with an incident.
        /// </summary>
        /// <param name="incidentNumber">Incident number.</param>
        /// <returns>Returns the list of workstream objects.</returns>
        [HttpGet]
        public async Task<IActionResult> GetAllWorkstremsAsync([FromQuery]string incidentNumber)
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
                        new Models.Error.Error
                        {
                            StatusCode = SignInErrorCode,
                            ErrorMessage = "Azure Active Directory access token for user is found empty.",
                        });
                }

                var workstreams = await this.workstreamStorageProvider.GetAllAsync(incidentNumber).ConfigureAwait(false);
                this.telemetryClient.TrackEvent($"Workstream data retrieved for {incidentNumber}");

                return this.Ok(workstreams);
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
