// <copyright file="IncidentListObject.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{
    /// <summary>
    /// Model to define data from both servicenow and tablestorage.
    /// </summary>
    public class IncidentListObject
    {
        /// <summary>
        /// Gets or sets conversation id of the user.
        /// </summary>
        public string PersonalConversationId { get; set; }

        /// <summary>
        /// Gets or sets conversation id of the message sent to the team.
        /// </summary>
        public string TeamConversationId { get; set; }

        /// <summary>
        /// Gets or sets reply to id.
        /// </summary>
        public string ReplyToId { get; set; }

        /// <summary>
        /// Gets or sets id of the message sent to the user.
        /// </summary>
        public string PersonalActivityId { get; set; }

        /// <summary>
        /// Gets or sets created on date.
        /// </summary>
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets updated on date.
        /// </summary>
        public string UpdatedOn { get; set; }

        /// <summary>
        /// Gets or sets prioriy of the ticket.
        /// </summary>
        public string Priority { get; set; }

        /// <summary>
        /// Gets or sets assigned to user displayName.
        /// </summary>
        public User AssignedTo { get; set; }

        /// <summary>
        /// Gets or sets id of the message sent to the team.
        /// </summary>
        public string TeamActivityId { get; set; }

        /// <summary>
        /// Gets or sets service url.
        /// </summary>
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets description of the incident.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets short description of the incident.
        /// </summary>
        public string ShortDescription { get; set; }

        /// <summary>
        /// Gets or sets scope of the incident.
        /// </summary>
        public string Scope { get; set; }

        /// <summary>
        /// Gets or sets status of the incident.
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// Gets or sets bridge id of the incident.
        /// </summary>
        public string BridgeId { get; set; }

        /// <summary>
        /// Gets or sets  bridge url of the incident.
        /// </summary>
        public string BridgeLink { get; set; }

        /// <summary>
        /// Gets or sets requested by display name.
        /// </summary>
        public User RequestedBy { get; set; }

        /// <summary>
        /// Gets or sets requested for display name.
        /// </summary>
        public User RequestedFor { get; set; }

        /// <summary>
        /// Gets or sets incident current activity.
        /// </summary>
        public string CurrentActivity { get; set; }

        /// <summary>
        /// Gets or sets sys_id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets incident number.
        /// </summary>
        public string Number { get; set; }
    }
}
