// <copyright file="StatusEntity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models.TableEntities
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Table used for storing user configuration.
    /// </summary>
    public class StatusEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StatusEntity"/> class.
        /// </summary>
        public StatusEntity()
        {
            this.PartitionKey = "status";
        }

        /// <summary>
        /// Gets or sets status title available.
        /// </summary>
        public string Status
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether status is currently used or not.
        /// </summary>
        public bool Active { get; set; }
    }
}
