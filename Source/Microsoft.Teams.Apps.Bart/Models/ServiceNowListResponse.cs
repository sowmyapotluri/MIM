// <copyright file="ServiceNowResponse.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing service now response details.
    /// </summary>
    public class ServiceNowListResponse
    {
        /// <summary>
        /// Gets or sets incident details.
        /// </summary>
        [JsonProperty("result")]
        public List<Incident> Incident { get; set; }
    }
}