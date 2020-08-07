// <copyright file="IncidentEntity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models.TableEntities
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Table used for storing additional incident.
    /// </summary>
    public class IncidentEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="IncidentEntity"/> class.
        /// </summary>
        public IncidentEntity()
        {
        }

        /// <summary>
        /// Gets or sets user Active Directory object Id.
        /// </summary>
        //public string UserAdObjectId
        //{
        //    get { return this.RowKey; }
        //    set { this.RowKey = value; }
        //}

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
        //public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets updated on date.
        /// </summary>
        //public string UpdatedOn { get; set; }

        /// <summary>
        /// Gets or sets prioriy of the ticket.
        /// </summary>
        public string Priority { get; set; }

        /// <summary>
        /// Gets or sets assigned to user displayName.
        /// </summary>
        public string AssignedTo { get; set; }

        /// <summary>
        /// Gets or sets assigned to user AadId.
        /// </summary>
        public string AssignedToId { get; set; }

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
        //public string Description { get; set; }

        /// <summary>
        /// Gets or sets short description of the incident.
        /// </summary>
        //public string ShortDescription { get; set; }

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
        public string RequestedBy { get; set; }

        /// <summary>
        /// Gets or sets requested by id.
        /// </summary>
        public string RequestedById { get; set; }

        /// <summary>
        /// Gets or sets requested for display name.
        /// </summary>
        public string RequestedFor { get; set; }

        /// <summary>
        /// Gets or sets requested for id.
        /// </summary>
        public string RequestedForId { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether request originated from technology support center.
        /// </summary>
        public string TSC { get; set; }
    }
}
