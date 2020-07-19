// <copyright file="User.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

using Newtonsoft.Json;

namespace Microsoft.Teams.Apps.Bart.Models
{
    /// <summary>
    /// Class containing data which are user specific.
    /// </summary>
    public class User
    {
        /// <summary>
        /// Gets or sets the id of user in AAD.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the upn of user in AAD.
        /// </summary>
        [JsonProperty("userPrincipalName")]
        public string UserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets the display name of user in AAD.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets the Guid in teams user id.
        /// </summary>
        [JsonProperty("teamsUserId")]
        public string TeamsUserId { get; set; }

        /// <summary>
        /// Gets or sets the service url.
        /// </summary>
        [JsonProperty("serviceUrl")]
        public string ServiceUrl { get; set; }
    }
}
