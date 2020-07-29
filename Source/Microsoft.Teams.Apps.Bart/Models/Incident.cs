// <copyright file="Incident.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing incident details.
    /// </summary>
    public class Incident
    {
        /// <summary>
        /// Gets or sets description.
        /// </summary>
        [JsonProperty("description")]
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets short description.
        /// </summary>
        [JsonProperty("short_description")]
        public string Short_Description { get; set; }

        /// <summary>
        /// Gets or sets status.
        /// </summary>
        [JsonProperty("u_status")]
        public string Status { get; set; }

        /// <summary>
        /// Gets or sets state.
        /// </summary>
        [JsonProperty("state")]
        public string State { get; set; }

        /// <summary>
        /// Gets or sets priority.
        /// </summary>
        [JsonProperty("priority")]
        public string Priority { get; set; }


        /// <summary>
        /// Gets or sets priority.
        /// </summary>
        [JsonProperty("severity")]
        public string Severity { get; set; }

        /// <summary>
        /// Gets or sets sys_id.
        /// </summary>
        [JsonProperty("sys_id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets incident number.
        /// </summary>
        [JsonProperty("number")]
        public string Number { get; set; }

        /// <summary>
        /// Gets or sets incident created datetime.
        /// </summary>
        [JsonProperty("sys_created_on")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets incident due datetime.
        /// </summary>
        [JsonProperty("due_date")]
        public string DueBy { get; set; }

        /// <summary>
        /// Gets or sets incident created datetime.
        /// </summary>
        [JsonProperty("u_current_activity")]
        public string CurrentActivity { get; set; }

        /// <summary>
        /// Gets or sets incident updated datetime.
        /// </summary>
        [JsonProperty("sys_updated_on")]
        public string UpdatedOn { get; set; }

        /// <summary>
        /// Gets or sets incident created datetime.
        /// </summary>
        [JsonProperty("work_notes")]
        public string WorkNotes { get; set; }

        /// <summary>
        /// Gets or sets incident urgency.
        /// </summary>
        [JsonProperty("urgency")]
        public string Urgency { get; set; }

        /// <summary>
        /// Gets or sets incident created datetime.
        /// </summary>
        [JsonProperty("impact")]
        public string Impact { get; set; }

        /// <summary>
        /// Gets or sets associated webex bridge code.
        /// </summary>
        [JsonProperty("bridge")]
        public string Bridge { get; set; }

        /// <summary>
        /// Gets or sets associated webex bridge.
        /// </summary>
        [JsonProperty("bridgeDetails")]
        public ConferenceRoomEntity BridgeDetails { get; set; }

        /// <summary>
        /// Gets or sets scope.
        /// </summary>
        [JsonProperty("scope")]
        public string Scope { get; set; }

        ///// <summary>
        ///// Gets or sets assignedTo display name.
        ///// </summary>
        //[JsonProperty("assignedTo")]
        //public string AssignedTo { get; set; }

        ///// <summary>
        ///// Gets or sets assignedTo id.
        ///// </summary>
        //[JsonProperty("assignedToId")]
        //public string AssignedToId { get; set; }

        /// <summary>
        /// Gets or sets requested by display name.
        /// </summary>
        [JsonProperty("requestedBy")]
        public string RequestedBy { get; set; }

        /// <summary>
        /// Gets or sets requested by id.
        /// </summary>
        [JsonProperty("requestedById")]
        public string RequestedById { get; set; }

        /// <summary>
        /// Gets or sets requested for display name.
        /// </summary>
        [JsonProperty("requestedFor")]
        public string RequestedFor { get; set; }

        /// <summary>
        /// Gets or sets requested for id.
        /// </summary>
        [JsonProperty("requestedForId")]
        public string RequestedForId { get; set; }

    }
}
