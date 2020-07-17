namespace Microsoft.Teams.Apps.Bart.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Teams.Apps.Bart.Helpers;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.Error;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Newtonsoft.Json;

    /// <summary>
    /// Exposes methods required for incident creation.
    /// </summary>
    public class ServiceNowProvider: IServiceNowProvider
    {
        /// <summary>
        /// Create incident API URL.
        /// </summary>
        private readonly string createIncident = "/api/now/table/incident";

        /// <summary>
        /// Update incident API URL.
        /// </summary>
        private readonly string updateIncident = "/api/now/table/incident/{0}";

        /// <summary>
        /// Users list URL.
        /// </summary>
        private readonly string searchUsers = "/api/now/table/sys_user?sysparm_query=name%3D{0}&sysparm_limit=10";

        /// <summary>
        /// Search incidents URL.
        /// </summary>
        private readonly string searchIncidents = "/api/now/table/incident?sysparm_query=short_descriptionLIKE{0}%5EORnumberLIKE{0}&sysparm_display_value=true&sysparm_fields=number,short_description,sys_created_on,work_notes,sys_id,state,sys_updated_on&sysparm_limit=10";

        /// <summary>
        /// Get all incidents URL.
        /// </summary>
        private readonly string allIncidents = "/api/now/table/incident?sysparm_query=short_descriptionLIKE{0}%5EORnumberLIKE{0}%5Esys_created_by%3DSVC_teams_automation&sysparm_fields=number,short_description,sys_created_on,work_notes,sys_id,state,sys_updated_on&sysparm_limit=10";

        /// <summary>
        /// Get new incidents URL.
        /// </summary>
        private readonly string newIncidents = "/api/now/table/incident?sysparm_query=short_descriptionLIKE{0}%5EORnumberLIKE{0}%5Estate%3D1%5Esys_created_by%3DSVC_teams_automation&sysparm_fields=number,short_description,sys_created_on,work_notes,sys_id,state,sys_updated_on&sysparm_limit=10";

        /// <summary>
        /// Get suspended incidents URL.
        /// </summary>
        private readonly string suspendedIncidents = "/api/now/table/incident?sysparm_query=short_descriptionLIKE{0}%5EORnumberLIKE{0}%5Estate%3D2%5Esys_created_by%3DSVC_teams_automation&sysparm_fields=number,short_description,sys_created_on,work_notes,sys_id,state,sys_updated_on&sysparm_limit=10";

        /// <summary>
        /// Get service restored incidents URL.
        /// </summary>
        private readonly string serviceRestoredIncidents = "/api/now/table/incident?sysparm_query=short_descriptionLIKE{0}%5EORnumberLIKE{0}%5Estate%3D3%5Esys_created_by%3DSVC_teams_automation&sysparm_fields=number,short_description,sys_created_on,work_notes,sys_id,state,sys_updated_on&sysparm_limit=10";

        /// <summary>
        /// Get recent incidents URL.
        /// </summary>
        private readonly string recentIncidents = "/api/now/table/incident?sysparm_query=short_descriptionLIKE{0}%5EORnumberLIKE{0}%5EORDERBYDESCsys_created_on%5Esys_created_by%3DSVC_teams_automation&sysparm_fields=number,short_description,sys_created_on,work_notes,sys_id,state,sys_updated_on&sysparm_limit=10";

        /// <summary>
        /// API helper service for making post and get calls to Graph.
        /// </summary>
        private readonly IApiHelper apiHelper;

        /// <summary>
        /// Telemetry client to log event and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceNowProvider"/> class.
        /// </summary>
        /// <param name="apiHelper">Api helper service for making post and get calls to Graph.</param>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        public ServiceNowProvider(IApiHelper apiHelper, TelemetryClient telemetryClient)
        {
            this.apiHelper = apiHelper;
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Create new incident.
        /// </summary>
        /// <param name="incident"><see cref="Incident"/> object. </param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Event response object.</returns>
        public async Task<dynamic> CreateIncidentAsync(Incident incident, string token)
        {
            var httpResponseMessage = await this.apiHelper.PostAsync(this.createIncident, token, JsonConvert.SerializeObject(incident, 
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore })).ConfigureAwait(false);
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<ServiceNowResponse>(content).Incident;
            }

            var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);

            this.telemetryClient.TrackTrace($"Create incident API failure- url: {this.createIncident}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
            var failureResponse = new
            {
                StatusCode = httpResponseMessage.StatusCode,
                ErrorResponse = errorResponse,
            };
            return failureResponse;
        }

        /// <summary>
        /// Update new incident.
        /// </summary>
        /// <param name="incident"><see cref="Incident"/> object. </param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Event response object.</returns>
        public async Task<dynamic> UpdateIncidentAsync(Incident incident, string token)
        {
            string payload = JsonConvert.SerializeObject(incident, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            var httpResponseMessage = await this.apiHelper.PatchAsync(string.Format(this.updateIncident, incident.Id), token, payload).ConfigureAwait(false);
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<ServiceNowResponse>(content).Incident;
            }

            var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);

            this.telemetryClient.TrackTrace($"Update incident API failure- url: {this.updateIncident}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
            var failureResponse = new
            {
                StatusCode = httpResponseMessage.StatusCode,
                ErrorResponse = errorResponse,
            };
            return failureResponse;
        }

        /// <summary>
        /// Get incident.
        /// </summary>
        /// <param name="incidentId">Incident id. </param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Event response object.</returns>
        public async Task<dynamic> GetIncidentAsync(string incidentId, string token)
        {
            var httpResponseMessage = await this.apiHelper.GetAsync(string.Format(this.updateIncident, incidentId), token).ConfigureAwait(false);
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<ServiceNowResponse>(content).Incident;
            }

            var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);

            this.telemetryClient.TrackTrace($"Update incident API failure- url: {this.updateIncident}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
            var failureResponse = new
            {
                StatusCode = httpResponseMessage.StatusCode,
                ErrorResponse = errorResponse,
            };
            return failureResponse;
        }

        /// <summary>
        /// Search users in service now.
        /// </summary>
        /// <param name="searchText">Search query from the user. </param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Event response object.</returns>
        public async Task<dynamic> SearchUsersAsync(string searchText, string token)
        {
            var httpResponseMessage = await this.apiHelper.GetAsync(string.Format(this.searchUsers, searchText), token).ConfigureAwait(false);
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<dynamic>(content);
            }

            var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);

            this.telemetryClient.TrackTrace($"Search users API failure- url: {this.searchUsers}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
            var failureResponse = new
            {
                StatusCode = httpResponseMessage.StatusCode,
                ErrorResponse = errorResponse,
            };
            return failureResponse;
        }

        /// <summary>
        /// Get incidents.
        /// </summary>
        /// <param name="searchQuery">Query for searching. </param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Event response object.</returns>
        public async Task<dynamic> SearchIncidentAsync(string searchQuery, string token)
        {
            var httpResponseMessage = await this.apiHelper.GetAsync(string.Format(this.searchIncidents, searchQuery), token).ConfigureAwait(false);
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<ServiceNowListResponse>(content).Incident;
            }

            var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);

            this.telemetryClient.TrackTrace($"Search incidents API failure- url: {this.searchUsers}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
            var failureResponse = new
            {
                StatusCode = httpResponseMessage.StatusCode,
                ErrorResponse = errorResponse,
            };
            return failureResponse;
        }

        /// <summary>
        /// Get incidents based on factors.
        /// </summary>
        /// <param name="commandId">Query for searching. </param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Event response object.</returns>
        public async Task<dynamic> GetIncidentAsync(string commandId, string searchQuery, string token)
        {
            string url = string.Format(this.recentIncidents, searchQuery);
            switch (commandId)
            {
                case "newincidents":
                    url = string.Format(this.newIncidents, searchQuery);
                    break;
                case "suspendedincidents":
                    url = string.Format(this.suspendedIncidents, searchQuery);
                    break;
                case "servicerestoredincidents":
                    url = string.Format(this.serviceRestoredIncidents, searchQuery);
                    break;
                case "allincidents":
                    url = string.Format(this.allIncidents, searchQuery);
                    break;
            }

            var httpResponseMessage = await this.apiHelper.GetAsync(url, token).ConfigureAwait(false);
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<ServiceNowListResponse>(content).Incident;
            }

            var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);

            this.telemetryClient.TrackTrace($"Search incidents API failure- url: {this.searchUsers}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
            var failureResponse = new
            {
                StatusCode = httpResponseMessage.StatusCode,
                ErrorResponse = errorResponse,
            };
            return failureResponse;
        }
    }
}
