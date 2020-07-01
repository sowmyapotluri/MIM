// <copyright file="IIncidentStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Providers.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on UserConfiguration table.
    /// </summary>
    public interface IIncidentStorageProvider
    {
        /// <summary>
        /// Add or update user configuration.
        /// </summary>
        /// <param name="incidentEntity">User configuration entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<bool> AddAsync(IncidentEntity incidentEntity);

        /// <summary>
        /// Get user configuration.
        /// </summary>
        /// <param name="incidentNumber">Active Directory object Id of user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<IncidentEntity> GetAsync(string partitionKey, string rowKey);
    }
}