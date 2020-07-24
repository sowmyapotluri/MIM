// <copyright file="AzureADOptions.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.Bart.Models
{
    /// <summary>
    ///  Options for configuring authentication using Azure Active Directory.
    /// </summary>
    public class AzureADOptions
    {
        /// <summary>
        /// Gets or sets the OpenID Connect authentication scheme to use for authentication with this instance of Azure Active Directory authentication.
        /// </summary>
        public string OpenIdConnectSchemeName { get; set; }

        /// <summary>
        /// Gets or sets the Cookie authentication scheme to use for sign in with this instance of Azure Active Directory authentication.
        /// </summary>
        public string CookieSchemeName { get; set; }

        /// <summary>
        /// Gets the Jwt bearer authentication scheme to use for validating access tokens for this instance of Azure Active Directory Bearer authentication.
        /// </summary>
        public string JwtBearerSchemeName { get; }

        /// <summary>
        /// Gets or sets the client Id.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Gets or sets the client secret.
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// Gets or sets the tenant Id.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the Azure Active Directory instance.
        /// </summary>
        public string Instance { get; set; }

        /// <summary>
        /// Gets or sets the domain of the Azure Active Directory tennant.
        /// </summary>
        public string Domain { get; set; }

        /// <summary>
        /// Gets or sets the sign in callback path.
        /// </summary>
        public string CallbackPath { get; set; }

        /// <summary>
        /// Gets or sets the sign out callback path.
        /// </summary>
        public string SignedOutCallbackPath { get; set; }

        /// <summary>
        /// Gets all the underlying authentication schemes.
        /// </summary>
        public string[] AllSchemes { get; }
    }
}
