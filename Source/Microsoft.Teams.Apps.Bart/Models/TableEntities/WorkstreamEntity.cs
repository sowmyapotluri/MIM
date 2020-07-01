// <copyright file="WorkstreamEntity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models.TableEntities
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Table used for storing additional incident.
    /// </summary>
    public class WorkstreamEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets initializes a new instance of the <see cref="IncidentEntity"/> class.
        /// </summary>
        //public IncidentEntity()
        //{
        //    this.PartitionKey = "msteams";
        //}

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
        public string AssignedTo { get; set; }

        /// <summary>
        /// Gets or sets windows time zone converted from IANA.
        /// </summary>
        public string Priority { get; set; }

        /// <summary>
        /// Gets or sets windows time zone converted from IANA.
        /// </summary>
        public string Description { get; set; }
    }
}
