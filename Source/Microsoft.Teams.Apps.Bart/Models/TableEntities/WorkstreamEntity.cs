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
        /// Initializes a new instance of the <see cref="WorkstreamEntity"/> class.
        /// </summary>
        /// <param name="incident">Incident entity.</param>
        public WorkstreamEntity(Incident incident)
        {
            this.PartitionKey = incident.Number;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WorkstreamEntity"/> class.
        /// </summary>
        public WorkstreamEntity()
        {
        }

        /// <summary>
        /// Gets or sets object Id.
        /// </summary>
        public string Id
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets displayName of user the workstream is assigned to.
        /// </summary>
        public string AssignedTo { get; set; }

        /// <summary>
        /// Gets or sets the priority of the workstream.
        /// </summary>
        public string Priority { get; set; }

        /// <summary>
        /// Gets or sets the description of the workstream.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether workstream is completed or not.
        /// </summary>
        public bool Status { get; set; }

        /// <summary>
        /// Gets or sets id of user the workstream is assigned to.
        /// </summary>
        public string AssignedToId { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether workstream is deleted or not.
        /// </summary>
        public bool InActive { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether workstream is recently created or not.
        /// </summary>
        public bool New { get; set; }
    }
}