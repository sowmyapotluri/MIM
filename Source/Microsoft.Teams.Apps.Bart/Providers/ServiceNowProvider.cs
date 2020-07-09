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
    }
}
