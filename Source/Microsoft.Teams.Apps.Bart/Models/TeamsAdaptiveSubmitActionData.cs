// <copyright file="TeamsAdaptiveSubmitActionData.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// Defines Teams-specific behavior for an adaptive card submit action.
    /// </summary>
    public class TeamsAdaptiveSubmitActionData
    {
        /// <summary>
        /// Gets or sets the Teams-specific action.
        /// </summary>
        [JsonProperty("msteams")]
        public CardAction MsTeams { get; set; }

        /// <summary>
        /// Gets or sets incident number.
        /// </summary>
        public string IncidentNumber { get; set; }

        /// <summary>
        /// Gets or sets incident id.
        /// </summary>
        public string IncidentId { get; set; }

        /// <summary>
        /// Gets or sets incident bridge number.
        /// </summary>
        public string BridgeId { get; set; }
    }
}
