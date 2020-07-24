// <copyright file="BatchRequestPayload.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{
    using System.Collections.Generic;

    /// <summary>
    /// creating <see cref="BatchRequestPayload"/> class.
    /// </summary>
    public class BatchRequestPayload
    {
        /// <summary>
        /// Gets or sets this method to get list of requests of each request in batch call.
        /// </summary>
        public List<dynamic> Requests { get; set; }
    }
}
