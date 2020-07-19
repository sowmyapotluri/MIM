// <copyright file="AdaptiveTaskModuleCardAction.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Defines model for opening task module.
    /// </summary>
    public class AdaptiveTaskModuleCardAction
    {
        /// <summary>
        /// Gets or sets action type for button.
        /// </summary>
        [JsonProperty("type")]
        public string Type
        {
            get
            {
                return "task/fetch";
            }
            set => this.Type = "task/fetch";
        }

        /// <summary>
        /// Gets or sets bot command to be used by bot for processing user inputs.
        /// </summary>
        [JsonProperty("text")]
        public string Text { get; set; }

        /// <summary>
        /// Gets or sets unique number to recognize incident.
        /// </summary>
        public string ActivityReferenceNumber { get; set; }

        /// <summary>
        /// Gets or sets unique GUID to recognize incident.
        /// </summary>
        public string ActivityReferenceId { get; set; }
    }
}
