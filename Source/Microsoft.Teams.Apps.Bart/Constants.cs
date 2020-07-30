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
        /// Text for take a tour action.
        /// </summary>
        public static readonly string TakeATour = "take a tour";

        /// <summary>
        /// Graph API base URL.
        /// </summary>
        public static readonly string GraphAPIBaseUrl = "https://graph.microsoft.com";

        /// <summary>
        /// Graph API for searching users URL.
        /// </summary>
        public static readonly string GraphApiToSearchUsers = "/v1.0/users?$filter=startswith(displayName,'{0}')&$select=displayName,userPrincipalName,id";

        /// <summary>
        /// Graph API to get team members URL.
        /// </summary>
        public static readonly string GraphApiToGetIncidemntManagers = "/v1.0/groups/{0}/members?$select=displayName,userPrincipalName,id";

        /// <summary>
        /// Graph API batch request URL.
        /// </summary>
        public static readonly string GraphBatchRequest = "https://graph.microsoft.com/v1.0/$batch";
    }
}
