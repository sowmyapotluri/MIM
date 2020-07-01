// <copyright file="ConferenceRoomEntity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models.TableEntities
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Table used for storing user configuration.
    /// </summary>
    public class ConferenceRoomEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConferenceRoomEntity"/> class.
        /// </summary>
        public ConferenceRoomEntity()
        {
            this.PartitionKey = "conferencerooms";
        }

        /// <summary>
        /// Gets or sets webex meeting access code.
        /// </summary>
        public string Code
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the webex room is available or not.
        /// </summary>
        public bool Available { get; set; }

        /// <summary>
        /// Gets or sets webex bridge url.
        /// </summary>
        public string BridgeURL { get; set; }

        /// <summary>
        /// Gets or sets teams channel id.
        /// </summary>
        public string ChannelId { get; set; }
    }
}
