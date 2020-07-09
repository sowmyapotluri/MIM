// <copyright file="ChangeTicketStatusPayload.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Represents the data payload of Action.Submit to change the status of a ticket.
    /// </summary>
    public class ChangeTicketStatusPayload
    {
        /// <summary>
        /// Action that set the status as new.
        /// </summary>
        public const string NewAction = "New";

        /// <summary>
        /// Action that set the status as suspended.
        /// </summary>
        public const string SuspendedAction = "Suspended";

        /// <summary>
        /// Action that set the status as service restored.
        /// </summary>
        public const string RestoredAction = "Service Restored";

        /// <summary>
        /// Gets or sets the incident id.
        /// </summary>
        [JsonProperty("incidentId")]
        public string IncidentId { get; set; }

        /// <summary>
        /// Gets or sets the incident number.
        /// </summary>
        [JsonProperty("incidentNumber")]
        public string IncidentNumber { get; set; }

        /// <summary>
        /// Gets or sets the status changes action to perform on the incident.
        /// </summary>
        [JsonProperty("action")]
        public string Action { get; set; }

        /// <summary>
        /// Gets or sets the title change action to perform on the incident.
        /// </summary>
        [JsonProperty("title")]
        public string Title { get; set; }
    }
}