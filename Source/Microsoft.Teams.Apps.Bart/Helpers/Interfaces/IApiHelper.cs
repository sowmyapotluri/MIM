// <copyright file="IApiHelper.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Helpers
{
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Threading.Tasks;

    /// <summary>
    /// Methods to perform API calls for GET, POST, PATCH requests.
    /// </summary>
    public interface IApiHelper
    {
        /// <summary>
        /// Method to perform HTTP GET requests in ServiceNow APIs.
        /// </summary>
        /// <typeparam name="T">Generic type class.</typeparam>
        /// <param name="url">Url to append on base Url for GET.(Example /api/messages).</param>
        /// <param name="token">Authentication token.</param>
        /// <param name="headers">Header parameters.</param>
        /// <returns>API response instance for GET request.</returns>
        Task<HttpResponseMessage> GetAsync(string url, string token, Dictionary<string, string> headers = null);

        /// <summary>
        /// Method to perform HTTP POST requests in ServiceNow APIs.
        /// </summary>
        /// <typeparam name="T">Generic Type class.</typeparam>
        /// <param name="url">Url to append on base Url for POST.(Example /api/messages).</param>
        /// <param name="token">Authentication token.</param>
        /// <param name="payload">request payload in JSON format.</param>
        /// <param name="headers">Header parameters.</param>
        /// <returns>API response instance for POST request.</returns>
        Task<HttpResponseMessage> PostAsync(string url, string token, string payload = "", Dictionary<string, string> headers = null);

        /// <summary>
        /// Method to perform HTTP PATCH requests in ServiceNow APIs.
        /// </summary>
        /// <typeparam name="T">Generic Type class.</typeparam>
        /// <param name="url">URL to append on base URL for POST.(Example /api/messages).</param>
        /// <param name="token">Authentication token.</param>
        /// <param name="payload">request payload in JSON format.</param>
        /// <param name="headers">Header parameters.</param>
        /// <returns>API response instance for POST request.</returns>
        Task<HttpResponseMessage> PatchAsync(string url, string token, string payload = "", Dictionary<string, string> headers = null);
    }
}