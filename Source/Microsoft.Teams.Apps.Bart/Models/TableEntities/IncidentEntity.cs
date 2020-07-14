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
        /// Gets or sets initializes a new instance of the <see cref="IncidentEntity"/> class.
        /// </summary>
        public IncidentEntity()
        {
            //this.PartitionKey = "msteams";
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
        /// Gets or sets id of the message sent to the user.
        /// </summary>
        public string PersonalConversationId { get; set; }

        /// <summary>
        /// Gets or sets windows time zone converted from IANA.
        /// </summary>
        public string TeamConversationId { get; set; }

        public string ReplyToId { get; set; }

        public string PersonalActivityId { get; set; }

        public string CreatedBy { get; set; }

        public string Priority { get; set; }

        public string AssignedTo { get; set; }

        public string AssignedToId { get; set; }

        public string TeamActivityId { get; set; }

        public string ServiceUrl { get; set; }

        public string Description { get; set; }

        public string ShortDescription { get; set; }

        public string Scope { get; set; }

        public string Status { get; set; }

        public string CreatedById { get; set; }

        public string BridgeId { get; set; }

        public string BridgeLink { get; set; }

    }
}
