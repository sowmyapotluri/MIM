// <copyright file="UserConfigurationEntity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models.TableEntities
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Table used for storing user configuration.
    /// </summary>
    public class UserConfigurationEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UserConfigurationEntity"/> class.
        /// </summary>
        public UserConfigurationEntity()
        {
            this.PartitionKey = "msteams";
        }

        /// <summary>
        /// Gets or sets user Active Directory object Id.
        /// </summary>
        public string UserAdObjectId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets conversation id for the user.
        /// </summary>
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or sets the Guid in teams user id.
        /// </summary>
        public string TeamsUserId { get; set; }

        /// <summary>
        /// Gets or sets the service url.
        /// </summary>
        public string ServiceUrl { get; set; }
    }
}
