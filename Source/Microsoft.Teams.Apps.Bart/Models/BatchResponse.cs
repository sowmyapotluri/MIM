// <copyright file="BatchResponse.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// creating <see cref="BatchResponse"/> class.
    /// </summary>
    /// <typeparam name="T">T type.</typeparam>
    public class BatchResponse<T>
    {
        /// <summary>
        /// Gets or sets this metod to ID.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets this metod to body.
        /// </summary>
        [JsonProperty("body")]
        public T Body { get; set; }
    }
}
