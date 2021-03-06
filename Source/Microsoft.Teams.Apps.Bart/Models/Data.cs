// <copyright file="Data.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{

    /// <summary>
    /// Class containing properties to be parsed from activity value.
    /// </summary>
    public class Data
    {
        /// <summary>
        /// Gets or sets bot command text.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Gets or sets unique number to recognize incident.
        /// </summary>
        public string ActivityReferenceNumber { get; set; }

        /// <summary>
        /// Gets or sets activity to update.
        /// </summary>
        public string Activity { get; set; }

        /// <summary>
        /// Gets or sets unique GUID to recognize incident.
        /// </summary>
        public string ActivityReferenceId { get; set; }
    }
}