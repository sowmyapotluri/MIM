// <copyright file="User.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

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
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the upn of user in AAD.
        /// </summary>
        public string UserPrincipleName { get; set; }

        /// <summary>
        /// Gets or sets the display name of user in AAD.
        /// </summary>
        public string DisplayName { get; set; }
    }
}
