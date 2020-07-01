// <copyright file="IStatusStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Providers.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on UserConfiguration table.
    /// </summary>
    public interface IStatusStorageProvider
    {
        /// <summary>
        /// Add or update user configuration.
        /// </summary>
        /// <param name="status">User configuration entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<bool> AddAsync(StatusEntity status);

        /// <summary>
        /// Get user configuration.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<List<StatusEntity>> GetAsync();
    }
}
