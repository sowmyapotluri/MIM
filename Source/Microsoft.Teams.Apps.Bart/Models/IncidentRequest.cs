// <copyright file="IncidentRequest.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;

    /// <summary>
    /// Incident request from user.
    /// </summary>
    public class IncidentRequest
    {
        /// <summary>
        /// Gets or sets incident details from the user.
        /// </summary>
        public Incident Incident { get; set; }

        /// <summary>
        /// Gets or sets list of workstream from the user.
        /// </summary>
        public List<WorkstreamEntity> Workstreams { get; set; }
    }
}
