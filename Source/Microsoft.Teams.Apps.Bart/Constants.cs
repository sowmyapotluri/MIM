// <copyright file="Constants.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;

    /// <summary>
    /// Constants class.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Graph API base URL.
        /// </summary>
        public static readonly string GraphAPIBaseUrl = "https://graph.microsoft.com";

        /// <summary>
        /// Graph API for getting users URL.
        /// </summary>
        public static readonly string SearchUsersGraphURL = "/v1.0/users?$filter=startswith(displayName,'{0}')";

        /// <summary>
        /// Text for take a tour action.
        /// </summary>
        public static readonly string TakeATour = "take a tour";

    }
}
