// <copyright file="IncidentApiController.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Connector;
    using Microsoft.Teams.Apps.Bart.Cards;
    using Microsoft.Teams.Apps.Bart.Helpers;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.Error;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.Teams.Apps.Bart.Providers.Storage;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using TimeZoneConverter;

    /// <summary>
    /// Incident API controller for handling API calls made from react js client app (used in task module).
    /// </summary>
    [ApiController]
    [Route("api/[controller]/[action]")]
    //[Authorize]
    public class IncidentApiController : ControllerBase
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
        /// Helper class which exposes methods required for incident creation and updation.
        /// </summary>
        private readonly IServiceNowProvider serviceNowProvider;

        /// <summary>
        /// Storage provider to perform creation and updation on Workstreans table.
        /// </summary>
        private readonly IWorkstreamStorageProvider workstreamStorageProvider;

        /// <summary>
        /// Storage provider to perform creation and updation on ConferenceBridge table.
        /// </summary>
        private readonly IConferenceBridgesStorageProvider conferenceBridgesStorageProvider;

        /// <summary>
        /// Storage provider to perform fetch operation on Incident table .
        /// </summary>
        private readonly IIncidentStorageProvider incidentStorageProvider;

        /// <summary>
        /// Graph API helper.
        /// </summary>
        private readonly IGraphApiHelper graphApiHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="IncidentApiController"/> class.
        /// </summary>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        /// <param name="tokenHelper">Generating and validating JWT token.</param>
        /// <param name="serviceNowProvider">Helper class which exposes methods required for incident creation.</param>
        /// <param name="conferenceBridgesStorageProvider">Helper class which exposes methods required for getting and updating conference room status.</param>
        /// <param name="incidentStorageProvider">Storage provider to perform fetch operation on Incident table.</param>
        /// <param name="workstreamStorageProvider">Storage provider to perform fetch operation on Workstream table.</param>
        /// <param name="graphApiHelper">Graph API helper.</param>
        public IncidentApiController(TelemetryClient telemetryClient, ITokenHelper tokenHelper, IServiceNowProvider serviceNowProvider, IConferenceBridgesStorageProvider conferenceBridgesStorageProvider, IIncidentStorageProvider incidentStorageProvider, IWorkstreamStorageProvider workstreamStorageProvider, IGraphApiHelper graphApiHelper)
        {
            this.telemetryClient = telemetryClient;
            this.tokenHelper = tokenHelper;
            this.serviceNowProvider = serviceNowProvider;
            this.conferenceBridgesStorageProvider = conferenceBridgesStorageProvider;
            this.incidentStorageProvider = incidentStorageProvider;
            this.workstreamStorageProvider = workstreamStorageProvider;
            this.graphApiHelper = graphApiHelper;
        }

        /// <summary>
        /// Create incident in ServiceNow.
        /// </summary>
        /// <param name="incidentRequest">Incident object.</param>
        /// <returns>Returns the newly created incident data.</returns>
        [HttpPost]
        public async Task<IActionResult> CreateIncidentAsync([FromBody]IncidentRequest incidentRequest)
        {
            try
            {
                Incident incident = incidentRequest.Incident; //JsonConvert.DeserializeObject<Incident>(incidentRequest.Incident.ToString());
                List<WorkstreamEntity> workstreams = incidentRequest.Workstreams; //JsonConvert.DeserializeObject<List<WorkstreamEntity>>(incidentRequest.Workstreams.ToString());
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

                var bridgeStatus = await this.conferenceBridgesStorageProvider.GetAsync(incident.Bridge).ConfigureAwait(false);
                if (bridgeStatus.Available)
                {
                    Incident incidentCreated = await this.serviceNowProvider.CreateIncidentAsync(incident, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa");
                    var incidentTableEntry = new IncidentEntity
                    {
                        PartitionKey = incidentCreated.Number,
                        RowKey = incidentCreated.Id,
                        BridgeId = incident.BridgeDetails.Code,
                        BridgeLink = incident.BridgeDetails.Code == "0" ? string.Empty : incident.BridgeDetails.BridgeURL,
                        RequestedBy = incident.RequestedBy,
                        RequestedById = incident.RequestedById,
                        RequestedFor = incident.RequestedFor,
                        RequestedForId = incident.RequestedForId,
                    };
                    await this.incidentStorageProvider.AddAsync(incidentTableEntry).ConfigureAwait(false);
                    //if (string.IsNullOrEmpty(incident.Id) && incident.Bridge == "0")
                    //{
                    //    bridgeStatus.Available = false;
                    //    await this.conferenceBridgesStorageProvider.AddAsync(bridgeStatus).ConfigureAwait(false);
                    //}
                    if (workstreams.Count > 0)
                    {
                        WorkstreamEntity workstreamEntity = new WorkstreamEntity(incidentCreated);
                        List<string> workstreamString = new List<string>();
                        foreach (var workstream in workstreams)
                        {
                            if (!string.IsNullOrEmpty(workstream.Description))
                            {
                                workstreamEntity.Id = Guid.NewGuid().ToString();
                                workstreamEntity.Description = workstream.Description;
                                workstreamEntity.AssignedTo = workstream.AssignedTo;
                                workstreamEntity.AssignedToId = workstream.AssignedToId;
                                workstreamEntity.Priority = workstream.Priority;
                                workstreamEntity.Status = workstream.Status;
                                workstreamEntity.New = workstream.New;

                                workstreamString.Add($"{workstreamEntity.Priority}: {workstream.Description}: {workstream.AssignedTo}: {workstream.Status}");
                                await this.workstreamStorageProvider.AddAsync(workstreamEntity).ConfigureAwait(false);
                            }
                        }

                        incidentCreated.WorkNotes = string.Join(',', workstreamString);
                        await this.serviceNowProvider.UpdateIncidentAsync(incidentCreated, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                    }

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
        /// Assign an incident to a user.
        /// </summary>
        /// <param name="assignedTo">Incident table entity object.</param>
        /// <returns>Returns a success status.</returns>
        [HttpPost]
        public async Task<IActionResult> AssignTicket([FromBody] IncidentEntity assignedTo)
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

                await this.incidentStorageProvider.AddAsync(assignedTo).ConfigureAwait(false);

                return this.Ok();

            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Get the assigned user to an incident from table storage.
        /// </summary>
        /// <param name="incidentNumber">Incident number.</param>
        /// <returns>Returns user object.</returns>
        [HttpGet]
        public async Task<IActionResult> AssignedUser([FromQuery] string incidentNumber)
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

                var incident = await this.incidentStorageProvider.GetAsync(incidentNumber).ConfigureAwait(false);
                User user = new User();
                if (!string.IsNullOrEmpty(incident.AssignedTo))
                {
                    user.DisplayName = incident.AssignedTo;
                    user.Id = incident.AssignedToId;
                }

                return this.Ok(user);

            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Get all incident from ServiceNow due this week.
        /// </summary>
        /// <param name="weekDay">Incident table entity object.</param>
        /// <returns>Returns the list of incidents.</returns>
        [HttpGet]
        public async Task<IActionResult> GetAllIncidents([FromQuery] string weekDay)
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

                var currentDay = Convert.ToDateTime(weekDay);
                var currentWeekDay = currentDay.DayOfWeek;
                int daysTillCurrentDay = currentWeekDay - DayOfWeek.Sunday;
                DateTime currentWeekStartDate = currentDay.AddDays(-daysTillCurrentDay).Date;
                DateTime currentWeekEndDate = currentDay.AddDays((double)(DayOfWeek.Saturday - daysTillCurrentDay)).Date;
                var query = string.Format(
                    "BETWEENjavascript%3Ags.dateGenerate('{0}-{1}-{2}'%2C'00%3A00%3A00')%40javascript%3Ags.dateGenerate('{3}-{4}-{5}'%2C'23%3A59%3A59')",
                    currentWeekStartDate.Year,
                    currentWeekStartDate.Month,
                    currentWeekStartDate.Day,
                    currentWeekEndDate.Year,
                    currentWeekEndDate.Month,
                    currentWeekEndDate.Day);
                List<IncidentListObject> incidentEntities = new List<IncidentListObject>();
                var incidents = await this.serviceNowProvider.SearchIncidentAsync(query, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                IncidentListObject listObject = new IncidentListObject();
                foreach (Incident incident in incidents)
                {
                    var incidentEntity = await this.incidentStorageProvider.GetAsync(incident.Number, incident.Id).ConfigureAwait(false);
                    if (incidentEntity != null)
                    {
                        string[] threadAndMessageId = incidentEntity.TeamConversationId.Split(";");
                        var threadId = threadAndMessageId[0];
                        var messageId = threadAndMessageId[1].Split("=")[1];

                        listObject.ShortDescription = incident.Short_Description;
                        listObject.Description = incident.Description;
                        listObject.CreatedOn = incident.CreatedOn;
                        listObject.UpdatedOn = incident.UpdatedOn;
                        listObject.Status = incidentEntity.Status;  // Till status options are figured out
                        listObject.CurrentActivity = incident.CurrentActivity;
                        listObject.Id = incident.Id;
                        listObject.Number = incident.Number;
                        listObject.TeamConversationId = $"https://teams.microsoft.com/l/message/{threadId}/{messageId}";
                        listObject.RequestedBy = new User
                        {
                            DisplayName = incidentEntity.RequestedBy == incidentEntity.RequestedFor ? incidentEntity.RequestedBy : incidentEntity.RequestedFor,
                            Id = incidentEntity.RequestedById == incidentEntity.RequestedById ? incidentEntity.RequestedById : incidentEntity.RequestedForId,
                        };
                        listObject.AssignedTo = new User
                        {
                            DisplayName = incidentEntity.AssignedTo,
                            Id = incidentEntity.AssignedToId,
                        };
                        incidentEntities.Add(listObject);
                    }
                }

                if (incidentEntities.Count > 0)
                {
                    var allRequests = new BatchRequestCreator().CreateBatchRequestPayloadForDetails(incidentEntities);
                    BatchRequestPayload payload = new BatchRequestPayload()
                    {
                        Requests = allRequests,
                    };

                    var result = await this.graphApiHelper.PostAsync("https://graph.microsoft.com/v1.0/$batch", token, JsonConvert.SerializeObject(payload));
                    var responseMessage = await result.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!string.IsNullOrEmpty(responseMessage))
                    {
                        var list = JObject.Parse(responseMessage)["responses"].ToObject<List<BatchResponse<dynamic>>>();
                        foreach (var response in list)
                        {
                            foreach (IncidentListObject incident in incidentEntities)
                            {
                                if (incident.RequestedBy.Id == response.Id)
                                {
                                    incident.RequestedBy.ProfilePicture = Convert.ToString(response.Body);
                                }

                                if (incident.AssignedTo.Id == response.Id)
                                {
                                    incident.AssignedTo.ProfilePicture = Convert.ToString(response.Body);
                                }
                            }
                        }
                    }
                }

                return this.Ok(incidentEntities);

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